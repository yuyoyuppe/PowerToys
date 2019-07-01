#include "pch.h"
#include "lowlevel_keyboard_event.h"
#include "powertoy_module.h"

namespace {
  HHOOK hook_handle = nullptr;
  HHOOK hook_handle_copy = nullptr; // make sure we do use nullptr in CallNextHookEx call
  LRESULT CALLBACK hook_proc(int nCode, WPARAM wParam, LPARAM lParam) {
    LowlevelKeyboardEvent event;
    if (nCode == HC_ACTION) {
      event.lParam = reinterpret_cast<KBDLLHOOKSTRUCT*>(lParam);
      event.wParam = wParam;
      auto supress = powertoys_events().signal_event(L"ll_keyboard", reinterpret_cast<intptr_t>(&event));
      return supress;
    } else {
      return CallNextHookEx(hook_handle_copy, nCode, wParam, lParam);
    }
  }
}

void start_lowlevel_keyboard_hook() {
  if (!hook_handle) {
    hook_handle = SetWindowsHookEx(WH_KEYBOARD_LL, hook_proc, GetModuleHandle(NULL), NULL);
    hook_handle_copy = hook_handle;
    if (!hook_handle) {
      throw std::runtime_error("Cannot install keyboard listener");
    }
  }
}

void stop_lowlevel_keyboard_hook() {
  if (hook_handle) {
    UnhookWindowsHookEx(hook_handle);
    hook_handle = nullptr;
  }
}
