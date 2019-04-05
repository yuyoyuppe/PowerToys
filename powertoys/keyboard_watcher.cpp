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
  LRESULT CALLBACK hook_proc(int nCode, WPARAM wParam, LPARAM lParam) {
    auto kb_hook = reinterpret_cast<KBDLLHOOKSTRUCT*>(lParam);
    if (nCode == HC_ACTION) {
      std::unique_lock<std::mutex> lock(hook_mutex);
      if (kb_hook->vkCode == VK_LWIN || kb_hook->vkCode == VK_RWIN) {
        if (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN) {
          winkey_pressed = true;
          winkey_press_timestamp = stdclock::now();
          lock.unlock();
          hook_cv.notify_one();
        } else {
          winkey_pressed = false;
          if (winkey_signaled) {
            lock.unlock();
            hook_cv.notify_one();
          }
        }
      } else if (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN) {
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
