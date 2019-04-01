#include "functionalities.h"
#include "popup_window.h"
#include "mouse_watcher.h"
#include "move_window.h"
namespace {
  PopupWindow* maximize_poup = NULL;

  RECT on_mouse_in(HWND hwnd, RECT buttons) {
    maximize_poup->show(buttons.left,
                        buttons.bottom,
                        buttons.right - buttons.left,
                        buttons.bottom - buttons.top);
    RECT result;
    result.left = buttons.left;
    result.top = buttons.bottom;
    result.right = buttons.right;
    result.top = 2 * buttons.bottom - buttons.top;
    move_window::to_screen_left(hwnd);
    return result;
  }

  void on_mouse_out() {
    maximize_poup->hide();
  }
}

void start_mouse_handler() {
  if (maximize_poup == NULL) {
    maximize_poup = new PopupWindow();
    start_mouse_watcher(300, on_mouse_in, on_mouse_out);
  }
}
