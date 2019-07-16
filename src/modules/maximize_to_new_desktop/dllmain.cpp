#include "pch.h"
#include <interface/powertoy_module_interface.h>
#include <interface/lowlevel_keyboard_event_data.h>
#include "window_manager_popup.h"
#include "mouse_watcher.h"
#include "trace.h"
#include <cpprest/json.h>
#include <common/settings_helpers.h>

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
    json::value _settings = json::value::object();
    _settings.as_object()[L"name"] = json::value::string(get_name());
    _settings.as_object()[L"description"] = json::value::string(L"Adds popup that Maximizes a Window to a new Desktop.");
    {
      json::value _properties = json::value::object(true); //Keep order
      {
        json::value _property = json::value::object();
        _property.as_object()[L"display_name"] = json::value::string(L"Remove a virtual desktop when the last window is restored");
        _property.as_object()[L"editor_type"] = json::value::string(L"bool_toggle");
        _property.as_object()[L"value"] = json::value::boolean(should_close_desktop_after_restoring_last_window);
        _properties.as_object()[L"close desktop on restore"] = _property;
      }
      _settings.as_object()[L"properties"] = _properties;
    }
    std::wstringstream ss;
    ss << _settings;
    std::wstring w_str = ss.str();
    const wchar_t* result_wstr = w_str.c_str();
    size_t result_len = wcslen(result_wstr)+1;
    wchar_t *result = new wchar_t[result_len];
    wcscpy_s(result, result_len, result_wstr);
    *config = result;
    return true;
  }

  virtual void free_get_config(const wchar_t* config) override {
    delete[] config;
  };
  virtual void set_config(const wchar_t* config) override {
    web::json::value j = web::json::value::parse(config);
    if (!j.is_object()) {
      // Should be an object.
      return;
    }
    if (!j.has_object_field(L"properties")) {
      // Should have a properties field.
      return;
    }
    web::json::value object_properties = j.at(L"properties");
    if (object_properties.has_object_field(L"close desktop on restore")) {
      web::json::value close_desktop = object_properties.at(L"close desktop on restore");
      if (close_desktop.has_boolean_field(L"value")) {
        should_close_desktop_after_restoring_last_window = close_desktop.at(L"value").as_bool();
        if (maximize_popup != nullptr) {
          maximize_popup->set_delete_after_restore(should_close_desktop_after_restoring_last_window);
        }
      }
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
    json::value _settings = json::value::object();
    std::wstring name = get_name();
    _settings.as_object()[L"name"] = json::value::string(name);
    {
      json::value _properties = json::value::object(true); //Keep order
      {
        json::value _property = json::value::object();
        _property.as_object()[L"value"] = json::value::boolean(should_close_desktop_after_restoring_last_window);
        _properties.as_object()[L"close desktop on restore"] = _property;
      }
      _settings.as_object()[L"properties"] = _properties;
    }
    PowerToysSettings::save_powertoy_settings_json(name, _settings);
  }
  catch (std::exception ex) {
    //Couldn't save the settings.
  }
}

void MTNDPowertoy::init_settings() {
  try {
    std::wstring name = this->get_name();
    json::value settings = PowerToysSettings::load_powertoy_settings_json(name);
    web::json::value object_properties = settings.at(L"properties");
    if (object_properties.has_object_field(L"close desktop on restore")) {
      web::json::value close_desktop = object_properties.at(L"close desktop on restore");
      if (close_desktop.has_boolean_field(L"value")) {
        should_close_desktop_after_restoring_last_window = close_desktop.at(L"value").as_bool();
      }
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
