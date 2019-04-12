#include "pch.h"
#include "functionalities.h"
#include "d2d_window.h"
#include "keyboard_watcher.h"
#include "monitors.h"
#include <sstream>

namespace {
  D2DWindow *winkey_popup;
  HWND desktop, shell;
  void on_held() {
    auto window = GetForegroundWindow();
    auto parent = GetParent(window);


    window = GetAncestor(window, GA_ROOT);
    if (window == desktop || window == shell)
      window = nullptr;
    LONG window_styles = window ? GetWindowLong(window, GWL_STYLE) : 0;
    if ((window_styles & WS_CHILD) || (window_styles & WS_DISABLED) || (window_styles & WS_POPUP))
      window = nullptr;
    winkey_popup->show(window);
  }

  void on_held_pressed(DWORD vkCode) {

  }

  void on_release() {
    winkey_popup->hide();
  }
}

void start_winkey_handler() {
  if (winkey_popup)
    return;
  desktop = GetDesktopWindow();
  shell = GetShellWindow();
  winkey_popup = new D2DWindow();
  start_winkey_watcher(300, on_held, on_held_pressed, on_release);
}
