#include "mouse_watcher.h"
#include <thread>
#include <chrono>
#include <mutex>
#include <Windows.h>
namespace {
  using stdclock = std::chrono::system_clock;

  bool mouse_initialized = false;
  MouseInProc mouse_in_cb;
  MouseOutProc mouse_out_cb;
  bool mousein_reset = true;
  stdclock::time_point mousein_timestamp;
  stdclock::duration mousein_wait;

  void mouse_thread_proc() {
    using namespace std::chrono_literals;
    while (true) {
      std::this_thread::sleep_for(100ms);
      POINT mouse_pos;
      GetCursorPos(&mouse_pos);
      auto mouse_window = WindowFromPoint(mouse_pos);
      if (mouse_window == NULL) {
        mouse_out_cb();
        continue;
      }
      mouse_window = GetAncestor(mouse_window, GA_ROOTOWNER);
      if (GetWindowLong(mouse_window, GWL_STYLE) & WS_CHILD) {
        mouse_out_cb();
        continue;
      }
      RECT window_rect;
      GetWindowRect(mouse_window, &window_rect);
      auto dpi = GetDpiForWindow(mouse_window);
      int buttons_width = 185 * dpi / 120;
      int buttons_height = 37 * dpi / 120;
      if (mouse_pos.x > window_rect.right - buttons_width && 
          mouse_pos.y < window_rect.top + buttons_height) {
        if (mousein_reset) {
          mousein_reset = false;
          mousein_timestamp = stdclock::now();
        } else if (stdclock::now() - mousein_timestamp > mousein_wait) {
          window_rect.left = window_rect.right - buttons_width;
          window_rect.bottom = window_rect.top + buttons_height;
          mouse_in_cb(mouse_window, window_rect);
        }
      } else {
        mousein_reset = true;
        mouse_out_cb();
      }
    }
  }
}

void start_mouse_watcher(int ms_delay, MouseInProc on_mouse_in, MouseOutProc on_mouse_out) {
  if (!mouse_initialized) {
    mouse_initialized = true;
    mouse_in_cb = on_mouse_in;
    mouse_out_cb = on_mouse_out;
    mousein_wait = std::chrono::milliseconds(ms_delay);
    std::thread(mouse_thread_proc).detach();
  }
}
