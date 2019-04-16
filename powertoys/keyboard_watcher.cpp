#include "pch.h"
#include "keyboard_watcher.h"
#include "start_visible.h"

namespace {
  using stdclock = std::chrono::system_clock;

  HHOOK hook_handle = NULL;
  std::mutex hook_mutex;
  std::condition_variable hook_cv;
  stdclock::time_point winkey_press_timestamp;
  bool winkey_pressed = false;
  bool winkey_signaled = false;
  std::function<void()> on_held_cb, on_relese_cb;
  std::function<void(unsigned long)> on_held_pressed_cb;

/*
  Uses SetWindowsHookEx to install system-wide hook that intercepts keyboard
  events. After WinKey is pressed, conditional variable is signaled and
  held_delay_thread_proc thread waits for specified amount of time. If the key
  is still pressed on_held callback is called.

  Takes care not to call any of the callbacks more than once for each event.
*/
  bool other_key_was_pressed = false;
  LRESULT CALLBACK hook_proc(int nCode, WPARAM wParam, LPARAM lParam) {
    auto kb_hook = reinterpret_cast<KBDLLHOOKSTRUCT*>(lParam);
    if (nCode == HC_ACTION) {
      std::unique_lock<std::mutex> lock(hook_mutex);
      if (kb_hook->vkCode == VK_LWIN || kb_hook->vkCode == VK_RWIN) {
        if (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN) {
          // Check if any other key is held
          BYTE keys_state[256];
          bool other_key_held = false;
          memset(keys_state, 0, 256);
          GetKeyboardState(keys_state);
          for (int vk = 0; vk < 256 && !other_key_held; ++vk) {
            if (vk == VK_LWIN || vk == VK_RWIN)
              continue;
            other_key_held = keys_state[vk] & 0x80; // test high bit
            // Pressing WinKey + M can get M key stuck in "pressed" state
            if (other_key_held)
              other_key_held = GetAsyncKeyState(vk) & 0x8000;
          }
          if (!other_key_held) {
            winkey_pressed = true;
            other_key_was_pressed = false;
            winkey_press_timestamp = stdclock::now();
            lock.unlock();
            hook_cv.notify_one();
          }
        } else { 
          winkey_pressed = false;
          if (winkey_signaled) {
            lock.unlock();
            hook_cv.notify_one();
            if (!other_key_was_pressed) {
              INPUT input[3] = { {}, {}, {} };
              input[0].type = INPUT_KEYBOARD;
              input[0].ki.wVk = VK_ESCAPE;
              input[1].type = INPUT_KEYBOARD;
              input[1].ki.wVk = VK_ESCAPE;
              input[1].ki.dwFlags = KEYEVENTF_KEYUP;
              input[2].type = INPUT_KEYBOARD;
              input[2].ki.wVk = VK_LWIN;
              input[2].ki.dwFlags = KEYEVENTF_KEYUP;
              SendInput(3, input, sizeof(INPUT));
              return 1;
            }
          }
        }
      } else if (winkey_signaled && (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN)) {
        if (kb_hook->vkCode == VK_OEM_COMMA) {
          // Special case - on comma hide our window.
          // We need to hide the window from this thread before OS does that for us.
          winkey_pressed = false;
          winkey_signaled = false;
          lock.unlock();
          on_relese_cb();
        } else {
          other_key_was_pressed = true;
        }
        on_held_pressed_cb(kb_hook->vkCode);
      }
    }
    return CallNextHookEx(hook_handle, nCode, wParam, lParam);
  }

  void held_delay_thread_proc(int ms) {
    auto delay = std::chrono::milliseconds(ms);
    while (true) {
      std::unique_lock<std::mutex> lock(hook_mutex);
      hook_cv.wait(lock, [] { return (winkey_pressed && !winkey_signaled) || (!winkey_pressed && winkey_signaled); });
      if (winkey_pressed) {
        auto wait_time = stdclock::now() - winkey_press_timestamp;
        while (winkey_pressed && wait_time <= delay) {
          lock.unlock();
          std::this_thread::sleep_for(delay - wait_time);
          lock.lock();
          wait_time = stdclock::now() - winkey_press_timestamp;
        }
        // Make sure not to call the callback if start menu is visible
        if (winkey_pressed && !is_start_visible()) {
          winkey_signaled = true;
          on_held_cb();
        }
      } else if (winkey_signaled) {
        winkey_signaled = false;
        lock.unlock();
        on_relese_cb();
      }
    }
  }
}
 
void start_winkey_watcher(int ms_delay, std::function<void()> on_held, std::function<void(unsigned long)> on_held_pressed, std::function<void()> on_released) {
  if (hook_handle == NULL) {
    on_held_cb = on_held;
    on_held_pressed_cb = on_held_pressed;
    on_relese_cb = on_released;
    winkey_pressed = false;
    winkey_signaled = false;
    std::thread(held_delay_thread_proc, ms_delay).detach();
    hook_handle = SetWindowsHookEx(WH_KEYBOARD_LL, hook_proc, GetModuleHandle(NULL), NULL);
    if (hook_handle == NULL) {
      throw std::runtime_error("Cannot install keyboard listener");
    }
  }
}
