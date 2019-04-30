#pragma once
#include<functional>

/*
  Initializes WinKey hook.

  When WinKey is held for at least for ms_delay milliseconds, the on_held
  callback will be called. It will not be called though, if the Start menu is
  visible.

  If while WinKey is held and another key is pressed, it will be reported through
  on_held_ pressed callback with virtual key code passed as DWORD parameter.

  When the WinKey is finally released, on_released callback will be called.
 */

void start_winkey_watcher(int ms_delay,
                          std::function<void()> on_held,
                          std::function<void(unsigned long)> on_held_pressed,
                          std::function<void()> on_released);
void signal_hide();
bool winkey_held();
bool only_winkey_key_held();
