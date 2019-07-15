#include "pch.h"
#include <interface/powertoy_module_interface.h>
#include <interface/lowlevel_keyboard_event_data.h>
#include "trace.h"
#include <cpprest/json.h>

using namespace web;

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
private:
  bool test_bool_prop = false;
  int test_int_prop = 3;
  std::wstring test_string_prop = L"a string";
  std::wstring test_color_prop = L"#1212FF";
  bool _enabled = false;
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
  // Return JSON with the configuration options.
  virtual bool get_config(const wchar_t** config) override {
    json::value _settings = json::value::object();
    _settings.as_object()[L"name"] = json::value::string(get_name());
    _settings.as_object()[L"description"] = json::value::string(L"Shows the different controls for the settings.");
    {
      json::value _properties = json::value::object(true); //Keep order
      {
        json::value _property = json::value::object();
        _property.as_object()[L"display_name"] = json::value::string(L"This is what a bool_toggle looks like");
        _property.as_object()[L"editor_type"] = json::value::string(L"bool_toggle");
        _property.as_object()[L"value"] = json::value::boolean(test_bool_prop);
        _properties.as_object()[L"test bool_toggle"] = _property;
      }
      {
        json::value _property = json::value::object();
        _property.as_object()[L"display_name"] = json::value::string(L"This is what a int_spinner looks like");
        _property.as_object()[L"editor_type"] = json::value::string(L"int_spinner");
        _property.as_object()[L"value"] = json::value::number(test_int_prop);
        _properties.as_object()[L"test int_spinner"] = _property;
      }
      {
        json::value _property = json::value::object();
        _property.as_object()[L"display_name"] = json::value::string(L"This is what a string_text looks like");
        _property.as_object()[L"editor_type"] = json::value::string(L"string_text");
        _property.as_object()[L"value"] = json::value::string(test_string_prop);
        _properties.as_object()[L"test string_text"] = _property;
      }
      {
        json::value _property = json::value::object();
        _property.as_object()[L"display_name"] = json::value::string(L"This is what a color_picker looks like");
        _property.as_object()[L"editor_type"] = json::value::string(L"color_picker");
        _property.as_object()[L"value"] = json::value::string(test_color_prop);
        _properties.as_object()[L"test color_picker"] = _property;
      }
      _settings.as_object()[L"properties"] = _properties;
    }
    std::wstringstream ss;
    ss << _settings;
    std::wstring w_str = ss.str();
    const wchar_t* result_wstr = w_str.c_str();
    size_t result_len = wcslen(result_wstr) + 1;
    wchar_t *result = new wchar_t[result_len];
    wcscpy_s(result, result_len, result_wstr);
    *config = result;
    return true;
  }
  virtual void free_get_config(const wchar_t* config) override {
    delete[] config;
  };
  // Passes JSON with the configuration settings for the powertoy
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
    if (object_properties.has_object_field(L"test bool_toggle")) {
      web::json::value object_value = object_properties.at(L"test bool_toggle");
      if (object_value.has_boolean_field(L"value")) {
        test_bool_prop=(object_value.at(L"value").as_bool());
      }
    }
    if (object_properties.has_object_field(L"test int_spinner")) {
      web::json::value object_value = object_properties.at(L"test int_spinner");
      if (object_value.has_number_field(L"value")) {
        test_int_prop = (object_value.at(L"value").as_integer());
      }
    }
    if (object_properties.has_object_field(L"test string_text")) {
      web::json::value object_value = object_properties.at(L"test string_text");
      if (object_value.has_string_field(L"value")) {
        test_string_prop = (object_value.at(L"value").as_string());
      }
    }
    if (object_properties.has_object_field(L"test color_picker")) {
      web::json::value object_value = object_properties.at(L"test color_picker");
      if (object_value.has_string_field(L"value")) {
        test_color_prop = (object_value.at(L"value").as_string());
      }
    }
  }

  // Enable the powertoy
  virtual void enable() {
    _enabled = true; 
  }

  // Disable the powertoy
  virtual void disable() {
    _enabled = false; 
  }

  // Returns if the powertoys is enabled
  virtual bool is_enabled() override { 
    return _enabled; 
  }

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


