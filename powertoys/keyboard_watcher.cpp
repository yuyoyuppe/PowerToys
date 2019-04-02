#include "keyboard_watcher.h"
#include <algorithm>
#include <chrono>
#include <mutex>
#include <thread>
#include <functional>
#include <condition_variable>
#include <Windows.h>
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

  LRESULT CALLBACK hook_proc(int nCode, WPARAM wParam, LPARAM lParam) {
    auto kb_hook = reinterpret_cast<KBDLLHOOKSTRUCT*>(lParam);
    if (nCode == HC_ACTION && (kb_hook->vkCode == VK_LWIN || kb_hook->vkCode == VK_RWIN)) {
      std::unique_lock<std::mutex> lock(hook_mutex);
      if ((wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN)) {
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
 
void start_winkey_watcher(int ms_delay, std::function<void()> on_held, std::function<void()> on_released) {
  if (hook_handle == NULL) {
    on_held_cb = on_held;
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
