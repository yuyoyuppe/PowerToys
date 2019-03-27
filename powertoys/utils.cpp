#include "utils.h"
#include <dwmapi.h>
#pragma comment(lib, "dwmapi.lib")

std::optional<RECT> getWindowMaximizeButton(HWND hwnd) {
  RECT button;
  auto res = DwmGetWindowAttribute(hwnd, DWMWA_CAPTION_BUTTON_BOUNDS, &button, sizeof(RECT));
  if (res == S_OK)
    return button;
  else
    return {};
}
