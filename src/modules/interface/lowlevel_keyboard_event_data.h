#pragma once
#include <Windows.h>

struct LowlevelKeyboardEvent {
  KBDLLHOOKSTRUCT* lParam;
  WPARAM wParam;
};
