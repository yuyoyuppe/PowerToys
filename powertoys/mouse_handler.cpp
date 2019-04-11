#include "pch.h"
#include "functionalities.h"
#include "window.h"
#include "mouse_watcher.h"
#include "move_window.h"
#include "virtual_desktops.h"

HWND current_window = NULL;

namespace {
  Window* maximize_poup = NULL;
  RECT on_mouse_in(HWND hwnd, RECT buttons) {
    auto dpi = GetDpiForWindow(hwnd);
    int offset_x = (50 * dpi) / 120;
    int offset_y = (20 * dpi) / 120;
    RECT result;
    result.left = buttons.left - offset_x;
    result.top = buttons.bottom + offset_y;
    result.right = buttons.right + offset_x;
    result.bottom = buttons.bottom + 4 * offset_y;
    maximize_poup->set_transparency(0)
                  .round_corners((15 * dpi) / 120)
                  .show(result)
                  .fade_in();
  current_window = hwnd;
    return result;
  }

  void on_mouse_out() {
    maximize_poup->fade_out();
  }
}



void start_mouse_handler() {
  if (maximize_poup == NULL) {
    maximize_poup = new Window();
  maximize_poup->add_handler(
    WM_LBUTTONDOWN,
    [=](HWND hwnd, WPARAM wparam, LPARAM lparam) {
      move_window_to_new_desktop(current_window);
      return 0;
    }
  );
    start_mouse_watcher(300, 100, on_mouse_in, on_mouse_out);
  }
}
