// powertoys_hook.cpp : Defines the exported functions for the DLL application.
//
#include "stdafx.h"
#define MOVETONEWDESKTOPMSGSTR "POWERTOYS_MOVE_TO_NEW_DESKTOP"
#include "virtual_desktops.h"

#define MESSAGE_TIMEOUT 300
ULONGLONG nextMessageTime = 0;

extern "C" __declspec(dllexport) LRESULT intercept_move_window_message(int code, WPARAM wParam, LPARAM lParam) {
  if (code < 0) {
    // No further processing to be done.
    return CallNextHookEx(NULL, code, wParam, lParam);
  }
  if (nextMessageTime >= GetTickCount64()) {
    // Some applications can have interactions that will keep repeating messages in a loop after moving the Windows to the next Desktop.
    // This is a workaround to keep those controlled.
    // TODO: Investigate if MSG->time could be used for this instead.
    return CallNextHookEx(NULL, code, wParam, lParam);
  }

  UINT msgcode = RegisterWindowMessage(MOVETONEWDESKTOPMSGSTR);

  PMSG msg;
  msg = (PMSG)lParam;

  if (msg->message == msgcode && msg->hwnd != NULL ) {
    move_window_to_new_desktop_impl(msg->hwnd, (int)(msg->lParam));
    nextMessageTime = GetTickCount64() + MESSAGE_TIMEOUT;
    return 0;
  }
  return CallNextHookEx(NULL, code, wParam, lParam);
}

