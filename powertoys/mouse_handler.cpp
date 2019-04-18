#include "pch.h"
#include "functionalities.h"
#include "d2d_window_manager_popup.h"
#include "mouse_watcher.h"
#include "move_window.h"
#include "virtual_desktops.h"

namespace {
  D2DWindowManagerPopup* maximize_poup = NULL;
  RECT on_mouse_in(HWND hwnd, RECT buttons) {
    auto dpi = GetDpiForWindow(hwnd);
    int width = buttons.right-buttons.left;
    int height = width/2;
    RECT result;
    result.left = buttons.left;
    result.top = buttons.bottom;
    result.right = buttons.right;
    result.bottom = buttons.bottom + height;
    maximize_poup->show(hwnd, result);
    return result;
  }

  void on_mouse_out() {
    maximize_poup->hide();
  }
}



void start_mouse_handler() {
  if (maximize_poup == NULL) {
    maximize_poup = new D2DWindowManagerPopup();
    start_mouse_watcher(300, 100, on_mouse_in, on_mouse_out);
  }
}
