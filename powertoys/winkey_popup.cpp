#include <sstream>
#include <string>
#include "functionalities.h"
#include "popup_window.h"
#include "keyboard_watcher.h"
#include "utils.h"


  PopupWindow& winkey_popup_window() {
    static PopupWindow window;
    return window;
  }

  PopupWindow& close_popup_windw() {
    static PopupWindow window([](HWND hwnd) { return 0; });
    return window;
  }

  void on_held() {
    HWND active_window = GetForegroundWindow();
    if (active_window == NULL) {
      return;
    }
    auto dpi = GetDpiForWindow(active_window);
    auto max_button = get__maximize_button(active_window);
    RECT rect;
    if (GetWindowRect(active_window, &rect) == 0) {
      return;
    }
    winkey_popup_window().show(
      rect.left + (rect.right - rect.left) / 2,
      rect.top + 150,
      600, 30,
      [=](HWND hwnd) {
      PAINTSTRUCT ps;
      HDC hdc = BeginPaint(hwnd, &ps);
      std::stringstream stream;
      stream << "HWND: " << active_window
        << " DPI: " << dpi;
      if (max_button) {
        stream << " MaxButton: (" << max_button->top << ", " << max_button->left << ")x("
          << max_button->bottom << ", " << max_button->right << ")";
      }
      else {
        stream << " MaxButton: Unknow";
      }
      auto str = stream.str();
      TextOut(hdc, 10, 5, str.c_str(), str.length());
      EndPaint(hwnd, &ps);
      return 0;
    });
    if (max_button) {
   //   close_popup_windw().show(max_button->left, max_button->top, max_button->right - max_button->left, max_button->top - max_button->bottom);
    }
  }

  void on_release() {
    winkey_popup_window().hide();
  }



void start_winkey_popup() {
  start_winkey_watcher(300, on_held, on_release);
}
