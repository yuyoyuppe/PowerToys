#include "pch.h"
#include "functionalities.h"
#include "d2d_window_manager_popup.h"
#include "mouse_watcher.h"
#include "virtual_desktops.h"

namespace {
  D2DWindowManagerPopup* maximize_popup = NULL;
  RECT on_mouse_in(HWND hwnd, RECT buttons, RECT monitor) {
    auto dpi = GetDpiForWindow(hwnd);
    int width = 120;
    int height = width / 3;
    LONG midPoint=(buttons.left+buttons.right)/2;
    RECT result;
    result.left = midPoint - width/2;
    result.top = buttons.bottom;
    result.right = midPoint + width/2;
    result.bottom = buttons.bottom + height;
    // Make sure resulting rect is inside monitor.
    result = keep_rect_inside_rect(result, monitor);
    maximize_popup->show(hwnd, result);
    return result;
  }

  void on_mouse_out() {
    maximize_popup->hide();
  }
}



void start_mouse_handler() {
  if (maximize_popup == NULL) {
    maximize_popup = new D2DWindowManagerPopup();
    start_mouse_watcher(300, 100, on_mouse_in, on_mouse_out, maximize_popup->get_hwnd());
  }
}
