#include "utils.h"
#include <dwmapi.h>
#pragma comment(lib, "dwmapi.lib")

std::optional<RECT> get_maximize_button_pos(HWND hwnd) {
  RECT button, window;
  auto res = DwmGetWindowAttribute(hwnd, DWMWA_CAPTION_BUTTON_BOUNDS, &button, sizeof(RECT));
  if (res == S_OK) {
    return button;
  } else {
    return {};
  }
}

std::optional<RECT> get_window_pos(HWND hwnd) {
  RECT window;
  if (GetWindowRect(hwnd, &window) == 0) {
    return {};
  } else {
    return window;
  }
}

int run_message_loop() {
  MSG msg;
  while (GetMessage(&msg, NULL, 0, 0)) {
    TranslateMessage(&msg);
    DispatchMessage(&msg);
  }
  return static_cast<int>(msg.wParam);
}