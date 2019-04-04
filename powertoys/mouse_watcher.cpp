#include "mouse_watcher.h"
#include "utils.h"
#include <Windows.h>
#include <thread>
#include <chrono>
#include <mutex>

namespace {
  using stdclock = std::chrono::system_clock;

  bool mouse_initialized = false;
  bool mousein_signalled = false;
  bool mousein_reset = true;
  MouseInProc mouse_in_cb;
  MouseOutProc mouse_out_cb;
  stdclock::time_point mousein_timestamp;
  stdclock::duration mousein_wait, mousein_sleep;
  RECT popup_rect, buttons_rect, stored_window_rect;

  bool mouse_in_rect(POINT mouse_pos, RECT rect) {
    return mouse_pos.x >= rect.left && mouse_pos.x <= rect.right &&
           mouse_pos.y >= rect.top  && mouse_pos.y <= rect.bottom;
  };

  // Test if mouse is over maximize button, popup window or in the trapezoid
  // between top edges of button and popup window.
  bool mouse_in_bounds(POINT mouse_pos, RECT buttons_rect, RECT popup_rect) {
    // Test if mouse is in one of the rects
    if (mouse_in_rect(mouse_pos, buttons_rect) || mouse_in_rect(mouse_pos, popup_rect))
      return true;
   // Test if mouse is in the trapezoid
    int top = buttons_rect.top;
    int bottom = popup_rect.top;
    int heigh = bottom - top;
    if (mouse_pos.y < top || mouse_pos.y > bottom || heigh == 0)
      return false;
    int top_left = buttons_rect.left;
    int top_right = buttons_rect.right;
    int bottom_left = popup_rect.left;
    int bottom_right = popup_rect.right;
    double dleft = double(bottom_left - top_left) / heigh;
    double dright = double(bottom_right - top_right) / heigh;
    int dy = mouse_pos.y - top;
    int px_left = top_left + (int)(dleft * dy);
    int px_right = top_right + (int)(dright * dy);
    return mouse_pos.x >= px_left && mouse_pos.x <= px_right;
  }
  void maybe_mouse_out() {
    if (mousein_signalled) {
      mousein_signalled = false;
      mouse_out_cb();
    }
  }
  void mouse_thread_proc() {
    while (true) {
      std::this_thread::sleep_for(mousein_sleep);
      auto mouse_pos = get_mouse_pos();
      if (!mouse_pos) {
        maybe_mouse_out();
        continue;
      }
      auto mouse_window = WindowFromPoint(*mouse_pos);
      if (mouse_window == nullptr) {
        maybe_mouse_out();
        continue;
      }
      mouse_window = GetAncestor(mouse_window, GA_ROOTOWNER);
      if (GetWindowLong(mouse_window, GWL_STYLE) & WS_CHILD) {
        maybe_mouse_out();
        continue;
      }
      auto window_rect = get_window_pos(mouse_window);
      if (!window_rect) {
        maybe_mouse_out();
        continue;
      }
      if (mousein_signalled) {
        if (window_rect->left != stored_window_rect.left ||
            window_rect->top != stored_window_rect.top ||
            window_rect->bottom != stored_window_rect.bottom ||
            window_rect->right != stored_window_rect.right ||
            !mouse_in_bounds(*mouse_pos, buttons_rect, popup_rect)) {
          mousein_signalled = false;
          mouse_out_cb();
        }
        continue;
      }
      auto dpi = GetDpiForWindow(mouse_window);
      int buttons_width = 185 * dpi / 120;
      int buttons_height = 37 * dpi / 120;
      buttons_rect.left = window_rect->right - buttons_width;
      buttons_rect.top = window_rect->top;
      buttons_rect.right = window_rect->right;
      buttons_rect.bottom = window_rect->top + buttons_height;
      if (mouse_in_rect(*mouse_pos, buttons_rect)) {
        if (mousein_reset) {
          mousein_reset = false;
          mousein_timestamp = stdclock::now();
        }
        if (stdclock::now() - mousein_timestamp > mousein_wait) {
          mousein_signalled = true;
          stored_window_rect = *window_rect;
          popup_rect = mouse_in_cb(mouse_window, buttons_rect);
        }
      } else {
        mousein_reset = true;
      }
    }
  }

}

void start_mouse_watcher(int ms_delay, int probe_ms_delay, MouseInProc on_mouse_in, MouseOutProc on_mouse_out) {
  if (!mouse_initialized) {
    mouse_initialized = true;
    mouse_in_cb = on_mouse_in;
    mouse_out_cb = on_mouse_out;
    mousein_wait = std::chrono::milliseconds(ms_delay);
    mousein_sleep = std::chrono::milliseconds(probe_ms_delay);
    std::thread(mouse_thread_proc).detach();
  }
}
