#include "pch.h"
#include "shortcut_guide.h"
#include "target_state.h"
#include "trace.h"
#include <cpprest/json.h>
#include <common/settings_helpers.h>

using namespace web;

OverlayWindow* instance = nullptr;

OverlayWindow::OverlayWindow() {
  init_settings();
}

const wchar_t * OverlayWindow::get_name() {
  return L"Shortcut Guide";
}

const wchar_t ** OverlayWindow::get_events() {
  static const wchar_t* events[2] = { L"ll_keyboard", 0 };
  return events;
}

bool OverlayWindow::get_config(const wchar_t** config) {
  json::value _settings = json::value::object();
  _settings.as_object()[L"name"] = json::value::string(get_name());
  _settings.as_object()[L"description"] = json::value::string(L"Shows a help overlay with Windows shortcuts when the Windows key is pressed.");
  {
    json::value _properties = json::value::object(true); //Keep order
    {
      json::value _property = json::value::object();
      _property.as_object()[L"display_name"] = json::value::string(L"How long to press the Windows key before showing the Shortcut Guide (ms)");
      _property.as_object()[L"editor_type"] = json::value::string(L"int_spinner");
      _property.as_object()[L"value"] = current_delay_setting;
      _properties.as_object()[L"press time"] = _property;
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

void OverlayWindow::free_get_config(const wchar_t* config) {
  delete[] config;
};

void OverlayWindow::set_config(const wchar_t * config) {
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
  if (object_properties.has_object_field(L"press time")) {
    web::json::value object_value = object_properties.at(L"press time");
    if (object_value.has_number_field(L"value")) {
      //Will be replaced by actual property update.
      int press_delay_time = (object_value.at(L"value").as_integer());
      current_delay_setting=press_delay_time;
      if (target_state) {
        target_state->set_delay(press_delay_time);
      }
    }
  }
  save_settings();
}

void OverlayWindow::enable() {
  if (!_enabled) {
    winkey_popup = new D2DOverlayWindow();
    target_state = new TargetState(current_delay_setting);
    winkey_popup->initialize();
    desktop = GetDesktopWindow();
    shell = GetShellWindow();
  }
  _enabled = true;
}

void OverlayWindow::disable() {
  if (_enabled) {
    winkey_popup->hide();
    target_state->exit();
    int a = 0;
    delete target_state;
    delete winkey_popup;
    target_state = nullptr;
    winkey_popup = nullptr;
  }
  _enabled = false;
}

bool OverlayWindow::is_enabled() {
  return _enabled;
}

intptr_t OverlayWindow::signal_event(const wchar_t * name, intptr_t data) {
  if (_enabled && wcscmp(name, L"ll_keyboard") == 0) {
    auto& event = *(reinterpret_cast<LowlevelKeyboardEvent*>(data));
    if (event.wParam == WM_KEYDOWN ||
        event.wParam == WM_SYSKEYDOWN ||
        event.wParam == WM_KEYUP ||
        event.wParam == WM_SYSKEYUP) {
      bool supress = target_state->signal_event(event.lParam->vkCode, 
                                                event.wParam == WM_KEYDOWN || event.wParam == WM_SYSKEYDOWN);
      return supress ? 1 : 0;
    }
  }
  return 0;
}

void OverlayWindow::on_held() {
  auto active_window = GetForegroundWindow();
  active_window = GetAncestor(active_window, GA_ROOT);
  if (active_window == desktop || active_window == shell) {
    active_window = nullptr;
  }
  auto window_styles = active_window ? GetWindowLong(active_window, GWL_STYLE) : 0;
  if ((window_styles & WS_CHILD) || (window_styles & WS_DISABLED)) {
    active_window = nullptr;
  }
  char class_name[256] = "";
  GetClassNameA(active_window, class_name, 256);
  if (strcmp(class_name, "SysListView32") == 0 ||
    strcmp(class_name, "WorkerW") == 0 ||
    strcmp(class_name, "Shell_TrayWnd") == 0 ||
    strcmp(class_name, "Shell_SecondaryTrayWnd") == 0) {
    active_window = nullptr;
  }
  winkey_popup->show(active_window);
}

void OverlayWindow::on_held_press(DWORD vkCode) {
  winkey_popup->animate(vkCode);
}

void OverlayWindow::was_hidden() {
  target_state->was_hiden();
}

void OverlayWindow::destroy() {
  delete this;
  instance = nullptr;
}

void OverlayWindow::init_settings() {
  try {
    std::wstring name = this->get_name();
    json::value settings = PowerToysSettings::load_powertoy_settings_json(name);
    web::json::value object_properties = settings.at(L"properties");
    if (object_properties.has_object_field(L"press time")) {
      web::json::value object_value = object_properties.at(L"press time");
      if (object_value.has_number_field(L"value")) {
        current_delay_setting = (object_value.at(L"value").as_integer());
      }
    }
  }
  catch (std::exception ex) {
    // Error while loading from the settings file. Just let default values stay as they are.
  }
}

void OverlayWindow::save_settings() {
  try {
    json::value _settings = json::value::object();
    std::wstring name = get_name();
    _settings.as_object()[L"name"] = json::value::string(name);
    {
      json::value _properties = json::value::object(true); //Keep order
      {
        json::value _property = json::value::object();
        _property.as_object()[L"value"] = json::value::number(current_delay_setting);
        _properties.as_object()[L"press time"] = _property;
      }
      _settings.as_object()[L"properties"] = _properties;
    }
    PowerToysSettings::save_powertoy_settings_json(name, _settings);
  }
  catch (std::exception ex) {
    //Couldn't save the settings.
  }
}
