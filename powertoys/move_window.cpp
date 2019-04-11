#include "pch.h"
#include "move_window.h"
#include "monitors.h"

namespace move_window {
  void maximize(HWND hwnd) {
    auto current_monitor = get_window_monitor(hwnd);
    RECT new_pos = current_monitor.rect;
    // This isn't exactly maximized. Seems to leave a bit of a distance from the borders.
    SetWindowPos(hwnd, HWND_TOP, new_pos.left, new_pos.top, new_pos.right - new_pos.left, new_pos.bottom - new_pos.top, SWP_SHOWWINDOW);
  }
  void restore(HWND hwnd) {}
  void minimize(HWND hwnd) {}

  void snap_top_left(HWND hwnd) {}
  void snap_top(HWND hwnd) {}
  void snap_top_right(HWND hwnd) {}
  void snap_left(HWND hwnd) {}
  void snap_right(HWND hwnd) {}
  void snap_bottom_left(HWND hwnd) {}
  void snap_bottom(HWND hwnd) {}
  void snap_bottom_right(HWND hwnd) {}

  void to_screen_left(HWND hwnd) {
    auto monitors = get_monitors();
    auto current_monitor = get_window_monitor(hwnd);
    auto current = get_monitor_index(monitors, current_monitor);
    auto left = current - 1;
    if (left == -1)
      left = monitors.size() - 1;
    RECT window_pos;
    if (GetWindowRect(hwnd, &window_pos) == 0) {
      window_pos = current_monitor.rect;
    }
    auto new_pos = translate_monitors(window_pos, current_monitor, monitors[left]);
    //auto new_pos = window_pos;
    SetWindowPos(hwnd, HWND_TOP, new_pos.left, new_pos.top, new_pos.right - new_pos.left, new_pos.bottom - new_pos.top, SWP_NOREDRAW);
  }
  void to_screen_right(HWND hwnd) {
    auto monitors = get_monitors();
    auto current_monitor = get_window_monitor(hwnd);
    auto current = get_monitor_index(monitors, current_monitor);
    auto right = current + 1;
    if (right == monitors.size())
      right = 0;
    RECT window_pos;
    if (GetWindowRect(hwnd, &window_pos) == 0) {
      window_pos = current_monitor.rect;
    }
    auto new_pos = translate_monitors(window_pos, current_monitor, monitors[right]);
    SetWindowPos(hwnd, HWND_TOP, new_pos.left, new_pos.top, new_pos.right - new_pos.left, new_pos.bottom - new_pos.top, SWP_NOREDRAW);
  }

  void to_next_desktop(HWND hwnd) {}
  void to_prev_desktop(HWND hwnd) {}
}
