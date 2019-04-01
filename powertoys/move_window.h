#pragma once
#include <Windows.h>

namespace move_window {
  void maximize(HWND hwnd);
  void restore(HWND hwnd);
  void minimize(HWND hwnd);

  void snap_top_left(HWND hwnd);
  void snap_top(HWND hwnd);
  void snap_top_right(HWND hwnd);
  void snap_left(HWND hwnd);
  void snap_right(HWND hwnd);
  void snap_bottom_left(HWND hwnd);
  void snap_bottom(HWND hwnd);
  void snap_bottom_right(HWND hwnd);

  void to_screen_left(HWND hwnd);
  void to_screen_right(HWND hwnd);

  void to_next_desktop(HWND hwnd);
  void to_prev_desktop(HWND hwnd);
}

