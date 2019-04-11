#include "pch.h"
#include "utils.h"
#include <dwmapi.h>
#pragma comment(lib, "dwmapi.lib")
#include <strsafe.h>


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
/*
bool operator==(const RECT& lhs, const RECT& rhs) {
  auto lhs_tuple = std::make_tuple(lhs.left, lhs.right, lhs.top, lhs.bottom);
  auto rhs_tuple = std::make_tuple(rhs.left, rhs.right, rhs.top, rhs.bottom);
  return lhs_tuple == rhs_tuple;
}

bool operator!=(const RECT& lhs, const RECT& rhs) {
  return !(lhs == rhs);
}
*/
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

void ShowLastErrorMessage(LPTSTR lpszFunction, DWORD dw)
{
  // Retrieve the system error message for the error code

  LPVOID lpMsgBuf;
  LPVOID lpDisplayBuf;

  FormatMessage(
    FORMAT_MESSAGE_ALLOCATE_BUFFER |
    FORMAT_MESSAGE_FROM_SYSTEM |
    FORMAT_MESSAGE_IGNORE_INSERTS,
    NULL,
    dw,
    MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
    (LPTSTR)&lpMsgBuf,
    0, NULL);

  // Display the error message and exit the process

  lpDisplayBuf = (LPVOID)LocalAlloc(LMEM_ZEROINIT,
    (lstrlen((LPCTSTR)lpMsgBuf) + lstrlen((LPCTSTR)lpszFunction) + 40) * sizeof(TCHAR));
  StringCchPrintf((LPTSTR)lpDisplayBuf,
    LocalSize(lpDisplayBuf) / sizeof(TCHAR),
    TEXT("%s failed with error %d: %s"),
    lpszFunction, dw, lpMsgBuf);
  MessageBox(NULL, (LPCTSTR)lpDisplayBuf, TEXT("Error"), MB_OK);

  LocalFree(lpMsgBuf);
  LocalFree(lpDisplayBuf);
}
