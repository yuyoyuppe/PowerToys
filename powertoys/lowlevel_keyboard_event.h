#pragma once
#include <Windows.h>

struct LowlevelKeyboardEvent {
  KBDLLHOOKSTRUCT* lParam;
  WPARAM wParam;
};

void start_lowlevel_keyboard_hook();
void stop_lowlevel_keyboard_hook();
