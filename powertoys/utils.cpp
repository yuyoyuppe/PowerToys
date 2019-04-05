#include "pch.h"
#include "utils.h"
#include <dwmapi.h>
#pragma comment(lib, "dwmapi.lib")

std::optional<RECT> get_button_pos(HWND hwnd) {
  RECT button;
  if (DwmGetWindowAttribute(hwnd, DWMWA_CAPTION_BUTTON_BOUNDS, &button, sizeof(RECT)) == S_OK) {
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

std::optional<POINT> get_mouse_pos() {
  POINT point;
  if (GetCursorPos(&point) == 0) {
    return {};
  } else {
    return point;
  }
}

bool operator==(const RECT& lhs, const RECT& rhs) {
  auto lhs_tuple = std::make_tuple(lhs.left, lhs.right, lhs.top, lhs.bottom);
  auto rhs_tuple = std::make_tuple(rhs.left, rhs.right, rhs.top, rhs.bottom);
  return lhs_tuple == rhs_tuple;
}

bool operator!=(const RECT& lhs, const RECT& rhs) {
  return !(lhs == rhs);
}

bool operator<(const RECT& lhs, const RECT& rhs) {
  auto lhs_tuple = std::make_tuple(lhs.left, lhs.right, lhs.top, lhs.bottom);
  auto rhs_tuple = std::make_tuple(rhs.left, rhs.right, rhs.top, rhs.bottom); 
  return lhs_tuple < rhs_tuple;
}
int run_message_loop() {
  MSG msg;
  while (GetMessage(&msg, NULL, 0, 0)) {
    TranslateMessage(&msg);
    DispatchMessage(&msg);
  }
  return static_cast<int>(msg.wParam);
}