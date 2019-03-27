#include "utils.h"
#include <dwmapi.h>
#pragma comment(lib, "dwmapi.lib")

std::optional<RECT> get_maximize_button(HWND hwnd) {
  RECT button;
  auto res = DwmGetWindowAttribute(hwnd, DWMWA_CAPTION_BUTTON_BOUNDS, &button, sizeof(RECT));
  if (res == S_OK)
    return button;
  else
    return {};
}

int run_message_loop() {
  MSG msg;
  while (GetMessage(&msg, NULL, 0, 0)) {
    TranslateMessage(&msg);
    DispatchMessage(&msg);
  }
  return static_cast<int>(msg.wParam);
}