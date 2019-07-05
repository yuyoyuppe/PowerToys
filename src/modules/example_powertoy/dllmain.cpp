#include "pch.h"
#include <interface/powertoy_module_interface.h>
#include <interface/lowlevel_keyboard_event_data.h>
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

// All methods called by the main PowerToys app
class ExamplePowertoy : public PowertoyModuleIface {
public:
  // Return the display name of the powertoy, this will be cached
  virtual const wchar_t* get_name() override {
    return L"Example Powertoy";
  }
  // Return array of the names of all events that this powertoy listens for, with
  // nullptr as the last element of the array. Nullptr can also be retured for empty
  // list.
  // Right now there is only lowlevel keyboard hook event
  virtual const wchar_t** get_events() override {
    static const wchar_t* events[2] = { L"ll_keyboard",
                                        nullptr };
    return events;
  }
  // Return JSON with the configuration options, will be cached
  virtual const wchar_t* get_config() override {
    return  L"";
  }
  // Passes JSON with the configuration settings for the powertoy
  virtual void set_config(const wchar_t* config) override { }
  // Enable the powertoy
  virtual void enable() { }
  // Disable the powertoy
  virtual void disable() { }
  // Handle incoming event, data is event-specific
  virtual intptr_t signal_event(const wchar_t* name, intptr_t data)  override {
    if (wcscmp(name, L"ll_keyboard") == 0) {
      auto& event = *(reinterpret_cast<LowlevelKeyboardEvent*>(data));
      // Return 1 if the keypress is to be suppressed (not forwarded to Windows),
      // otherwise return 0.
      return 0;
    }
    return 0;
  }
  // Destroy the powertoy and free memory
  virtual void destroy() override {
    delete this;
  }
};

extern "C" __declspec(dllexport) PowertoyModuleIface*  __cdecl powertoy_create() {
  return new ExamplePowertoy();
}


