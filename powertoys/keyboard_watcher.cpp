#include "keyboard_watcher.h"
#include <algorithm>
#include <chrono>
#include <mutex>
#include <thread>
#include <functional>
#include <condition_variable>
#include <Windows.h>

namespace {
  using stdclock = std::chrono::system_clock;

  HHOOK kbhook_handle = NULL;
  std::mutex kbhook_mutex;
  std::condition_variable kbhook_cv;
  stdclock::time_point winkey_press_timestamp;
  bool winkey_pressed = false;
  bool winkey_signaled = false;
  std::function<void()> on_held_cb, on_relese_cb;

  LRESULT CALLBACK keyboard_hook(int nCode, WPARAM wParam, LPARAM lParam) {
    auto kb_hook = reinterpret_cast<KBDLLHOOKSTRUCT*>(lParam);
    if (nCode != HC_ACTION) {
      goto call_next_hook;
    }
    if (kb_hook->vkCode != VK_LWIN && kb_hook->vkCode != VK_RWIN) {
      goto call_next_hook;
    }
    if (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN) {
      std::unique_lock<std::mutex> lock(kbhook_mutex);
      winkey_pressed = true;
      winkey_press_timestamp = stdclock::now();
      lock.unlock();
      kbhook_cv.notify_one();
    } else {
      std::unique_lock<std::mutex> lock(kbhook_mutex);
      winkey_pressed = false;
      if (winkey_signaled) {
        winkey_signaled = false;
        on_relese_cb();
      }
    }
call_next_hook:
    // If we are in "signalled" state and keypress arrives - call on_held_cb
    // again, to update windows postion. 
    if (winkey_signaled) {
      std::thread([] {
        std::this_thread::sleep_for(std::chrono::milliseconds(100));
        if (winkey_signaled) {
          on_held_cb();
        }
      }).detach();
    }
    return CallNextHookEx(kbhook_handle, nCode, wParam, lParam);
  }

  void kbhook_delay_thread_proc(int ms) {
    auto delay = std::chrono::milliseconds(ms);
    while (true) {
      std::unique_lock<std::mutex> lock(kbhook_mutex);
      kbhook_cv.wait(lock, [] { return winkey_pressed && !winkey_signaled; });
      lock.unlock();
      auto wait_time = stdclock::now() - winkey_press_timestamp;
      while (wait_time <= delay) {
        std::this_thread::sleep_for(delay - wait_time);
        wait_time = stdclock::now() - winkey_press_timestamp;
      }
      lock.lock();
      if (winkey_pressed && !winkey_signaled) {
        winkey_signaled = true;
        on_held_cb();
      }
    }
  }
}
 
void start_winkey_watcher(int ms_delay, std::function<void()> on_held, std::function<void()> on_released) {
  std::lock_guard<std::mutex> lock(kbhook_mutex);
  if (kbhook_handle == NULL) {
    winkey_pressed = false;
    winkey_signaled = false;
    kbhook_handle = SetWindowsHookEx(WH_KEYBOARD_LL, keyboard_hook, GetModuleHandle(NULL), NULL);
    if (kbhook_handle == NULL) {
      throw std::runtime_error("Cannot install keyboard listener");
    }
    std::thread(kbhook_delay_thread_proc, ms_delay).detach();    
  }
  on_held_cb = on_held;
  on_relese_cb = on_released;
}
