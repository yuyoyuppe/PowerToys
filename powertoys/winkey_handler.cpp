#include "pch.h"
#include "functionalities.h"
#include "window.h"
#include "keyboard_watcher.h"
#include "monitors.h"
#include <sstream>

namespace {
  Window *winkey_popup;
  void on_held() {
    auto primary = get_primary_monitor();
    winkey_popup->show(
      primary.left(), primary.top(), primary.width(), primary.height() + 1,
      [=](HWND hwnd, WPARAM, LPARAM) {
      PAINTSTRUCT ps;
      HDC hdc = BeginPaint(hwnd, &ps);
      std::stringstream stream;
      stream << "Hullo!";
      std::string str = stream.str();
      TextOut(hdc, 20, 20, str.c_str(), str.length());
      EndPaint(hwnd, &ps);
      return 0;
    }).fade_in();
  }

  void on_held_pressed(DWORD vkCode) {

  }

  void on_release() {
    winkey_popup->fade_out();
  }
}

void start_winkey_handler() {
  if (winkey_popup)
    return;
  winkey_popup = new Window();
  start_winkey_watcher(300, on_held, on_held_pressed, on_release);
}
