#include "pch.h"
#include "shortcut_guide.h"
#include "target_state.h"
#include "trace.h"
#include <common/settings_objects.h>

OverlayWindow* instance = nullptr;

OverlayWindow::OverlayWindow() {
  init_settings();
}

const wchar_t * OverlayWindow::get_name() {
  return L"Shortcut Guide";
}

const wchar_t ** OverlayWindow::get_events() {
  static const wchar_t* events[2] = { ll_keyboard, 0 };
  return events;
}

bool OverlayWindow::get_config(const wchar_t** config) {
  PowerToysSettings::Settings _settings(
    get_name(),
    L"Shows a help overlay with Windows shortcuts when the Windows key is pressed."
  );

  PowerToysSettings::IntSpinnerPropertySetting delay_property(
    L"press time",
    L"How long to press the Windows key before showing the Shortcut Guide (ms)",
    current_delay_setting
  );
  delay_property.setSpinnerMax(10000);
  delay_property.setSpinnerMin(100);
  delay_property.setSpinnerStep(100);
  _settings.add_property(delay_property);

  *config = _settings.to_allocated_cstring();
  return true;
}

void OverlayWindow::free_get_config(const wchar_t* config) {
  PowerToysSettings::Settings::free_allocated_cstring(config);
};

void OverlayWindow::set_config(const wchar_t * config) {
  try {
    PowerToysSettings::PowerToyValues _values =
      PowerToysSettings::PowerToyValues::from_json_string(config);
    if (_values.is_int_value(L"press time")) {
      int press_delay_time = _values.get_int_value(L"press time");
      current_delay_setting = press_delay_time;
      if (target_state) {
        target_state->set_delay(press_delay_time);
      }
    }
  }
  catch (std::exception ex) {
    // Improper JSON.
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
  if (_enabled && wcscmp(name, ll_keyboard) == 0) {
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
    PowerToysSettings::PowerToyValues settings =
      PowerToysSettings::PowerToyValues::load_from_settings_file(get_name());
    if (settings.is_int_value(L"press time")) {
      current_delay_setting = settings.get_int_value(L"press time");
    }
  }
  catch (std::exception ex) {
    // Error while loading from the settings file. Just let default values stay as they are.
  }
}

void OverlayWindow::save_settings() {
  try {
    PowerToysSettings::PowerToyValues values(
      get_name()
    );
    values.add_property_value(
      PowerToysSettings::IntSettingsValue(
        L"press time",
        current_delay_setting
      )
    );
    values.save_to_settings_file();
  }
  catch (std::exception ex) {
    //Couldn't save the settings.
  }
}
