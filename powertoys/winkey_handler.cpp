#include "pch.h"
#include "functionalities.h"
#include "d2d_overlay_window.h"
#include "keyboard_watcher.h"
#include "monitors.h"
#include <sstream>

namespace {
  D2DOverlayWindow *winkey_popup;
  HWND desktop, shell;
  HWND active_window;
  void on_held() {
    auto active_window = GetForegroundWindow();
    active_window = GetAncestor(active_window, GA_ROOT);
    if (active_window == desktop || active_window == shell)
      active_window = nullptr;
    LONG window_styles = active_window ? GetWindowLong(active_window, GWL_STYLE) : 0;
    if ((window_styles & WS_CHILD) || (window_styles & WS_DISABLED))
      active_window = nullptr;
    char class_name[256] = "";
    GetClassNameA(active_window, class_name, 256);
    if (strcmp(class_name, "SysListView32") == 0 ||
        strcmp(class_name, "WorkerW") == 0 ||
        strcmp(class_name, "Shell_TrayWnd") == 0 ||
        strcmp(class_name, "Shell_SecondaryTrayWnd") == 0)
      active_window = nullptr;
    winkey_popup->show(active_window);
  }

  void on_held_pressed(DWORD vkCode) {
    winkey_popup->animate(vkCode);
  }

  void on_release() {
    //Calling hide here might cause a concurrency issue.
    //winkey_popup->hide();
  }
}

void start_winkey_handler() {
  if (winkey_popup)
    return;
  desktop = GetDesktopWindow();
  shell = GetShellWindow();
  winkey_popup = new D2DOverlayWindow();
  winkey_popup->initialize();
  start_winkey_watcher(300, on_held, on_held_pressed, on_release);
}
