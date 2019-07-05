#include "pch.h"
#include "keyboard_state.h"

bool winkey_held() {
  auto left = GetAsyncKeyState(VK_LWIN);
  auto right = GetAsyncKeyState(VK_RWIN);
  return (left & 0x8000) || (right & 0x8000);
}

bool only_winkey_key_held() {
  BYTE keys_state[256];
  memset(keys_state, 0, 256);
  GetKeyboardState(keys_state);
  for (int vk = 0; vk < 256; ++vk) {
    if (vk == VK_LWIN || vk == VK_RWIN)
      continue;
    auto key_held = keys_state[vk] & 0x80; // test high bit
    // Pressing WinKey + M can get M key stuck in "pressed" state
    if (key_held)
      key_held = GetAsyncKeyState(vk) & 0x8000;
    if (key_held)
      return false;
  }
  return true;
}