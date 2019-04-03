#include "functionalities.h"
#include "popup_window.h"
#include "mouse_watcher.h"
#include "move_window.h"
#include "utils.h"

namespace {
  PopupWindow* maximize_poup = NULL;
  
  RECT on_mouse_in(HWND hwnd, RECT buttons) {
    maximize_poup->set_transparency(0);
    maximize_poup->show(buttons.left - 50,
                        buttons.bottom + 20,
                        buttons.right - buttons.left + 100,
                        buttons.bottom - buttons.top);
    maximize_poup->fade_in();
    RECT result;
    result.left = buttons.left - 50;
    result.top = buttons.bottom + 20;
    result.right = buttons.right + 50;
    result.bottom = buttons.bottom + (buttons.bottom - buttons.top) + 20;
    return result;
  }

  void on_mouse_out() {
    maximize_poup->fade_out();
  }
}

void start_mouse_handler() {
  if (maximize_poup == NULL) {
    maximize_poup = new PopupWindow();
    start_mouse_watcher(300, 100, on_mouse_in, on_mouse_out);
  }
}
