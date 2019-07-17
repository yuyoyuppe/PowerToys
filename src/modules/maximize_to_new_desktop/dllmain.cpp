#include "pch.h"
#include <interface/powertoy_module_interface.h>
#include <interface/lowlevel_keyboard_event_data.h>
#include "window_manager_popup.h"
#include "mouse_watcher.h"
#include "trace.h"
#include <cpprest/json.h>
#include <common/settings_objects.h>

using namespace web;

//Forward declarations
RECT on_mouse_in(HWND hwnd, RECT buttons, RECT monitor);
void on_mouse_out();

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
  bool _enabled = false;
  virtual const wchar_t* get_name() override {
    return L"Move To New Desktop";
  }

  virtual const wchar_t** get_events() override {
    return nullptr;
  }

  virtual bool get_config(const wchar_t** config) override {
    PowerToysSettings::Settings _settings(
      get_name(),
      L"Adds popup that Maximizes a Window to a new Desktop."
    );
    _settings.add_property(
      PowerToysSettings::BoolTogglePropertySetting(
        L"close desktop on restore",
        L"Remove a virtual desktop when the last window is restored",
        should_close_desktop_after_restoring_last_window
      )
    );
    *config = _settings.to_allocated_cstring();
    return true;
  }

  virtual void free_get_config(const wchar_t* config) override {
    PowerToysSettings::Settings::free_allocated_cstring(config);
  };

  virtual void set_config(const wchar_t* config) override {
    try {
      PowerToysSettings::PowerToyValues _values =
        PowerToysSettings::PowerToyValues::from_json_string(config);
      if (_values.is_bool_value(L"close desktop on restore")) {
        should_close_desktop_after_restoring_last_window = _values.get_bool_value(L"close desktop on restore");
        if (maximize_popup != nullptr) {
          maximize_popup->set_delete_after_restore(should_close_desktop_after_restoring_last_window);
        }
      }
    }
    catch (std::exception ex) {
      // Improper JSON.
    }

    // Persist settings.
    save_settings();
  }
  virtual void enable() override {
    if (!_enabled) {
      maximize_popup = new D2DWindowManagerPopup();
      maximize_popup->set_delete_after_restore(should_close_desktop_after_restoring_last_window);
      start_mouse_watcher(400, 100, on_mouse_in, on_mouse_out, maximize_popup->get_hwnd());
    }
    _enabled = true;
  }
  virtual void disable() override {
    if (_enabled) {
      stop_mouse_watcher();
      delete maximize_popup;
    }
    _enabled = false;
  }
  virtual bool is_enabled() override {
    return _enabled;
  }
  virtual intptr_t signal_event(const wchar_t* name, intptr_t data)  override {
    return 0;
  }
  
  virtual void destroy() override {
    delete this;
    instance = nullptr;
  }
  static D2DWindowManagerPopup* maximize_popup;

private:
  bool should_close_desktop_after_restoring_last_window = true;
  void init_settings();
  void save_settings();
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

void MTNDPowertoy::save_settings() {
  try {
    PowerToysSettings::PowerToyValues values(
      get_name()
    );
    values.add_property_value(
      PowerToysSettings::BoolSettingsValue(
        L"close desktop on restore",
        should_close_desktop_after_restoring_last_window
      )
    );
    values.save_to_settings_file();
  }
  catch (std::exception ex) {
    //Couldn't save the settings.
  }
}

void MTNDPowertoy::init_settings() {
  try {
    PowerToysSettings::PowerToyValues settings =
      PowerToysSettings::PowerToyValues::load_from_settings_file(get_name());
    if (settings.is_bool_value(L"close desktop on restore")) {
      should_close_desktop_after_restoring_last_window = settings.get_bool_value(L"close desktop on restore");
    }
  }
  catch (std::exception ex) {
    // Error while loading from the settings file. Just let default values stay as they are.
  }
}

MTNDPowertoy::MTNDPowertoy() {
  // Initialize settings values
  init_settings();
}

extern "C" __declspec(dllexport) PowertoyModuleIface*  __cdecl powertoy_create() {
  if (!instance) {
    instance = new MTNDPowertoy();
    return instance;
  } else {
    return nullptr;
  }
}
