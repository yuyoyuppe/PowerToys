#include "pch.h"
#include <interface/powertoy_module_interface.h>
#include <interface/lowlevel_keyboard_event_data.h>
#include "window_manager_popup.h"
#include "mouse_watcher.h"
#include "trace.h"

BOOL APIENTRY DllMain(HMODULE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved) {
  switch (ul_reason_for_call) {
  case DLL_PROCESS_ATTACH:
    Trace::RegisterProvider();
    break;
  case DLL_THREAD_ATTACH:
  case DLL_THREAD_DETACH:
    break;
  case DLL_PROCESS_DETACH:
    Trace::UnregisterProvider();
    break;
  }
  return TRUE;
}

class MTNDPowertoy* instance = nullptr;

class MTNDPowertoy : public PowertoyModuleIface {
public:
  MTNDPowertoy();

  virtual const wchar_t* get_name() override {
    return L"Move To New Desktop Powertoy";
  }

  virtual const wchar_t** get_events() override {
    return nullptr;
  }

  virtual const wchar_t* get_config() override {
    return  L"";
  }
  virtual void set_config(const wchar_t* config) override { }
  virtual void enable() { }
  virtual void disable() { }
  virtual intptr_t signal_event(const wchar_t* name, intptr_t data)  override {
    return 0;
  }
  
  virtual void destroy() override {
    stop_mouse_watcher();
    delete maximize_popup;
    delete this;
    instance = nullptr;
  }
  static D2DWindowManagerPopup* maximize_popup;
};

D2DWindowManagerPopup* MTNDPowertoy::maximize_popup = nullptr;
std::chrono::steady_clock::time_point shown_start_time;

RECT on_mouse_in(HWND hwnd, RECT buttons, RECT monitor) {
  auto dpi = GetDpiForWindow(hwnd);
  int width = 120;
  int height = width / 3;
  LONG midPoint = (buttons.left + buttons.right) / 2;
  RECT result;
  result.left = midPoint - width / 2;
  result.top = buttons.bottom;
  result.right = midPoint + width / 2;
  result.bottom = buttons.bottom + height;
  // Make sure resulting rect is inside monitor.
  result = keep_rect_inside_rect(result, monitor);
  MTNDPowertoy::maximize_popup->show(hwnd, result);
  Trace::EventShow();
  shown_start_time = std::chrono::steady_clock::now();
  return result;
}

void on_mouse_out() {
  MTNDPowertoy::maximize_popup->hide();
  std::chrono::steady_clock::time_point shown_end_time = std::chrono::steady_clock::now();
  Trace::EventHide(std::chrono::duration_cast<std::chrono::milliseconds>(shown_end_time - shown_start_time).count());
}

MTNDPowertoy::MTNDPowertoy() {
  maximize_popup = new D2DWindowManagerPopup();
  start_mouse_watcher(400, 100, on_mouse_in, on_mouse_out, maximize_popup->get_hwnd());
}


extern "C" __declspec(dllexport) PowertoyModuleIface*  __cdecl powertoy_create() {
  if (!instance) {
    instance = new MTNDPowertoy();
    return instance;
  } else {
    return nullptr;
  }
}
