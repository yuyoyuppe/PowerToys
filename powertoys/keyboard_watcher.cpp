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

  HHOOK hook_handle = NULL;
  std::mutex held_delay_mutex, held_monitor_mutex;
  std::condition_variable held_delay_cv, held_monitor_cv;
  stdclock::time_point winkey_press_timestamp;
  bool winkey_pressed = false;
  bool winkey_signaled = false;
  std::function<void()> on_held_cb, on_relese_cb;

  LRESULT CALLBACK hook_proc(int nCode, WPARAM wParam, LPARAM lParam) {
    auto kb_hook = reinterpret_cast<KBDLLHOOKSTRUCT*>(lParam);
    if (nCode != HC_ACTION) {
      goto call_next_hook;
    }
    if (kb_hook->vkCode != VK_LWIN && kb_hook->vkCode != VK_RWIN) {
      goto call_next_hook;
    }
    if (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN) {
      std::unique_lock<std::mutex> lock(held_delay_mutex);
      winkey_pressed = true;
      winkey_press_timestamp = stdclock::now();
      lock.unlock();
      held_delay_cv.notify_one();
    } else {
      std::unique_lock<std::mutex> lock(held_delay_mutex);
      winkey_pressed = false;
      if (winkey_signaled) {
        winkey_signaled = false;
        on_relese_cb();
      }
    }
call_next_hook:
    return CallNextHookEx(hook_handle, nCode, wParam, lParam);
  }

  void held_delay_thread_proc(int ms) {
    auto delay = std::chrono::milliseconds(ms);
    while (true) {
      std::unique_lock<std::mutex> lock(held_delay_mutex);
      held_delay_cv.wait(lock, [] { return winkey_pressed && !winkey_signaled; });
      lock.unlock();
      auto wait_time = stdclock::now() - winkey_press_timestamp;
      while (wait_time <= delay) {
        std::this_thread::sleep_for(delay - wait_time);
        wait_time = stdclock::now() - winkey_press_timestamp;
      }
      lock.lock();
      if (winkey_pressed && !winkey_signaled) {
        winkey_signaled = true;
        held_monitor_cv.notify_one();
        on_held_cb();
      }
    }
  }

  void held_monitor_thread_proc() {
    while (true) {
      std::unique_lock<std::mutex> lock(held_monitor_mutex);
      held_monitor_cv.wait(lock, [] { return winkey_pressed && winkey_signaled; });
      std::this_thread::sleep_for(std::chrono::milliseconds(8));
      winkey_pressed = (GetKeyState(VK_LWIN) & 0x8000) || (GetKeyState(VK_RWIN) & 0x8000);
      if (winkey_signaled) {
        if (winkey_pressed) {
          on_held_cb();
        } else {
          winkey_signaled = false;
          on_relese_cb();
        }
      }
    }
  }
}
 
void start_winkey_watcher(int ms_delay, std::function<void()> on_held, std::function<void()> on_released) {
  std::lock_guard<std::mutex> lock(held_delay_mutex);
  if (hook_handle == NULL) {
    winkey_pressed = false;
    winkey_signaled = false;
    hook_handle = SetWindowsHookEx(WH_KEYBOARD_LL, hook_proc, GetModuleHandle(NULL), NULL);
    if (hook_handle == NULL) {
      throw std::runtime_error("Cannot install keyboard listener");
    }
    std::thread(held_delay_thread_proc, ms_delay).detach();
    std::thread(held_monitor_thread_proc).detach();
  }
  on_held_cb = on_held;
  on_relese_cb = on_released;
}
