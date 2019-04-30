#include "pch.h"
#include "keyboard_watcher.h"
#include "start_visible.h"
#include <deque>

namespace {
  struct KeyEvent {
    bool key_down;
    unsigned vk_code;
  };
  class TargetState {
  public:
    TargetState(std::function<void()> on_held, int ms_delay,
                std::function<void()> on_relese,
                std::function<void(unsigned long)> on_held_pressed_cb) :
      on_held_cb(on_held), delay(ms_delay),
      on_relese_cb(on_relese),
      on_held_pressed_cb(on_held_pressed_cb),
      thread(&TargetState::thread_proc, this)
    { }

    bool signal(unsigned vk_code, bool key_down) {
      std::unique_lock<std::mutex> lock(mutex);
      // some special case:
      if (state == Shown && key_down &&
          (vk_code == VK_OEM_COMMA ||
           vk_code == 0x4C || // L
           vk_code == 0x54 || // T
           (vk_code >= 0x30 && vk_code <= 0x39))) {
        state = Hidden;
        on_relese_cb();
        return false;
      }
      if (!events.empty() && events.back().key_down == key_down && events.back().vk_code == vk_code)
        return false;
      bool supress = false;
      if (!key_down && (vk_code == VK_LWIN || vk_code == VK_RWIN) &&
          state == Shown &&
          std::chrono::system_clock::now() - singnal_timestamp > std::chrono::seconds(1) &&
          !key_was_pressed) {
        supress = true;
      }
      events.push_back({ key_down, vk_code });
      lock.unlock();
      cv.notify_one();
      if (supress) {
        INPUT input[3] = { {}, {}, {} };
        input[0].type = INPUT_KEYBOARD;
        input[0].ki.wVk = VK_CONTROL;
        input[1].type = INPUT_KEYBOARD;
        input[1].ki.wVk = VK_CONTROL;
        input[1].ki.dwFlags = KEYEVENTF_KEYUP;
        input[2].type = INPUT_KEYBOARD;
        input[2].ki.wVk = VK_LWIN;
        input[2].ki.dwFlags = KEYEVENTF_KEYUP;
        SendInput(3, input, sizeof(INPUT));
      }
      return supress;
    }

    void was_hiden() {
      std::lock_guard<std::mutex> lock(mutex);
      state = Hidden;
      events.clear();
    }

    void exit() {  
      std::lock_guard<std::mutex> lock(mutex);
      events.clear();
      state = Exiting;
      cv.notify_one();
      thread.join();
    }
  private:
    KeyEvent next() {
      auto e = events.front();
      events.pop_front();
      return e;
    }

    void handle_hidden() {
      std::unique_lock<std::mutex> lock(mutex);
      if (events.empty())
        cv.wait(lock);
      if (events.empty() || state == Exiting)
        return;
      auto event = next();
      if (event.key_down && (event.vk_code == VK_LWIN || event.vk_code == VK_RWIN)) {
        state = Timeout;
        winkey_timestamp = std::chrono::system_clock::now();
      }
    }

    void handle_timeout() {
      std::unique_lock<std::mutex> lock(mutex);
      auto wait_time = delay - (std::chrono::system_clock::now() - winkey_timestamp);
      if (events.empty())
        cv.wait_for(lock, delay);
      if (state == Exiting)
        return;
      while (!events.empty()) {
        auto event = events.front();
        if (event.key_down && (event.vk_code == VK_LWIN || event.vk_code == VK_RWIN))
          events.pop_front();
        else
          break;
      }
      if (!events.empty() || !only_winkey_key_held() || is_start_visible()) {
        state = Hidden;
        return;
      }
      if (std::chrono::system_clock::now() - winkey_timestamp < delay)
        return;
      singnal_timestamp = std::chrono::system_clock::now();
      key_was_pressed = false;
      state = Shown;
      lock.unlock();
      on_held_cb();
    }

    void handle_shown() {
      std::unique_lock<std::mutex> lock(mutex);
      if (events.empty())
        cv.wait(lock);
      if (events.empty() || state == Exiting)
        return;
      auto event = next();
      if (event.key_down && (event.vk_code == VK_LWIN || event.vk_code == VK_RWIN))
        return;
      if (!event.key_down && (event.vk_code == VK_LWIN || event.vk_code == VK_RWIN) || !winkey_held()) {
        state = Hidden;
        lock.unlock();
        on_relese_cb();
        return;
      }
      if (event.key_down) {
        key_was_pressed = true;
        lock.unlock();
        on_held_pressed_cb(event.vk_code);
      }
    }

    void thread_proc() {
      while (true) {
        switch (state) {
        case Hidden:
          handle_hidden();
          break;
        case Timeout:
          handle_timeout();
          break;
        case Shown:
          handle_shown();
          break;
        case Exiting:
        default:
          return;
        }
      }
    }
    std::mutex mutex;
    std::condition_variable cv;
    std::chrono::system_clock::time_point winkey_timestamp, singnal_timestamp;
    std::chrono::milliseconds delay;
    std::function<void()> on_held_cb, on_relese_cb;
    std::function<void(unsigned long)> on_held_pressed_cb;
    std::deque<KeyEvent> events;
    enum { Hidden, Timeout, Shown, Exiting } state = Hidden;
    bool key_was_pressed = false;
    std::thread thread;
  };

  HHOOK hook_handle = NULL;
  TargetState* target_state;

  LRESULT CALLBACK hook_proc(int nCode, WPARAM wParam, LPARAM lParam) {
    auto kb_hook = reinterpret_cast<KBDLLHOOKSTRUCT*>(lParam);
    if (nCode == HC_ACTION &&
        (wParam == WM_KEYDOWN ||
         wParam == WM_SYSKEYDOWN ||
         wParam == WM_KEYUP ||
         wParam == WM_SYSKEYUP)) {
      bool supress = target_state->signal(kb_hook->vkCode, wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN);
      return supress ? 1 : 0;
    } else {
      return CallNextHookEx(hook_handle, nCode, wParam, lParam);
    }
  }
}

void start_winkey_watcher(int ms_delay, std::function<void()> on_held, std::function<void(unsigned long)> on_held_pressed, std::function<void()> on_released) {
  if (hook_handle == NULL) {
    target_state = new TargetState(on_held, ms_delay, on_released, on_held_pressed);
    hook_handle = SetWindowsHookEx(WH_KEYBOARD_LL, hook_proc, GetModuleHandle(NULL), NULL);
    if (hook_handle == NULL) {
      throw std::runtime_error("Cannot install keyboard listener");
    }
  }
}

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

void signal_hide() {
  if (target_state)
    target_state->was_hiden();
}
