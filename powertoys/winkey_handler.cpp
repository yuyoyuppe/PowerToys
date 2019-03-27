#include <sstream>
#include <string>
#include "functionalities.h"
#include "popup_window.h"
#include "keyboard_watcher.h"
#include "utils.h"

namespace {
  PopupWindow *winkey_popup;
  void on_held() {
    HWND active_window = GetForegroundWindow();
    if (active_window == NULL) {
      return;
    }

    auto dpi = GetDpiForWindow(active_window);
    auto window = get_window_pos(active_window);
    if (!window) {
      return;
    }
    winkey_popup->show(
      window->left,
      window->top,
      window->right - window->left,
      window->bottom - window->top,
      [=](HWND hwnd) {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        std::stringstream stream;
        stream << "HWND: " << active_window
               << " DPI: " << dpi
               << " Window: (" << window->left << "," << window->top << ")x("
                               << window->right << "," << window->bottom << ")";
        auto max_button = get_maximize_button_pos(active_window);
        if (max_button) {
          stream << " Button: " << max_button->right - max_button->left << ", " << max_button->bottom - max_button->top;
        }
        auto str = stream.str();
        TextOut(hdc, 20, 20, str.c_str(), str.length());
        EndPaint(hwnd, &ps);
        return 0;
      });
    SetForegroundWindow(active_window);
  }

  void on_release() {
    winkey_popup->hide();
  }
}

void start_winkey_handler() {
  if (winkey_popup)
    return;
  winkey_popup = new PopupWindow();
  start_winkey_watcher(300, on_held, on_release);
}
