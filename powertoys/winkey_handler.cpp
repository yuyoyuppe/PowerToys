#include <sstream>
#include <string>
#include "functionalities.h"
#include "popup_window.h"
#include "keyboard_watcher.h"
#include "monitors.h"
#include "utils.h"

namespace {
  PopupWindow *winkey_popup;
  void on_held() {
    winkey_popup->set_transparency(0);
    std::thread([] {
      SetForegroundWindow(winkey_popup->hwnd);
      auto primary = get_primary_monitor();
      winkey_popup->show(
        primary.left(), primary.top(), primary.width(), primary.height()+1,
        [=](HWND hwnd) {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        std::string str = "content here";
        TextOut(hdc, 20, 20, str.c_str(), str.length());
        EndPaint(hwnd, &ps);
        return 0;
      });
      for (double alpha = 0; alpha < 0.7; alpha += 0.08) {
        winkey_popup->set_transparency(alpha);
        std::this_thread::sleep_for(std::chrono::milliseconds(16));
      }
    }).detach();
  }

  void on_release() {
    std::thread([] {
      for (double alpha = 0.7; alpha > 0; alpha -= 0.08) {
        winkey_popup->set_transparency(alpha);
        std::this_thread::sleep_for(std::chrono::milliseconds(16));
      }
      winkey_popup->hide();
    }).detach();
  }
}

void start_winkey_handler() {
  if (winkey_popup)
    return;
  winkey_popup = new PopupWindow();
  start_winkey_watcher(300, on_held, on_release);
}
