#include "pch.h"
#include <interface/powertoy_module_interface.h>
#include <interface/lowlevel_keyboard_event_data.h>
#include "window_manager_popup.h"
#include "mouse_watcher.h"
#include "trace.h"
#include <common/settings_objects.h>
#include <ShellScalingApi.h>
#include <common/monitors.h>
#include <common/dpi_aware.h>

extern "C" IMAGE_DOS_HEADER __ImageBase;

//Forward declarations
RECT on_mouse_in(HWND hwnd, RECT buttons, POINT mouse_pos);
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

  virtual bool get_config(wchar_t* buffer, int *buffer_size) override {
    PowerToysSettings::Settings settings(
      get_name(),
      L"Adds a popup to maximize a window to a new desktop."
    );

    settings.set_icon_key(L"pt-maximize-new-desktop");
    
    HINSTANCE hinstance = reinterpret_cast<HINSTANCE>(&__ImageBase);
    wchar_t description[256];
    LoadString(hinstance, closeDesktopOnRestoreLastWindow.resourceId, description, ARRAYSIZE(description));
    settings.add_property(
      PowerToysSettings::BoolTogglePropertySetting(
        closeDesktopOnRestoreLastWindow.name,
        description,
        closeDesktopOnRestoreLastWindow.value
      )
    );

    LoadString(hinstance, popupDelay.resourceId, description, ARRAYSIZE(description));
    PowerToysSettings::IntSpinnerPropertySetting delay_property(
      popupDelay.name,
      description,
      popupDelay.value
    );
    delay_property.setSpinnerMax(10000);
    delay_property.setSpinnerMin(100);
    delay_property.setSpinnerStep(100);
    settings.add_property(delay_property);

    return settings.serialize_to_buffer(buffer, buffer_size);
  }

  virtual void set_config(const wchar_t* config) override {
    try {
      PowerToysSettings::PowerToyValues _values =
        PowerToysSettings::PowerToyValues::from_json_string(config);
      if (_values.is_bool_value(closeDesktopOnRestoreLastWindow.name)) {
        closeDesktopOnRestoreLastWindow.value = _values.get_bool_value(closeDesktopOnRestoreLastWindow.name);
        if (maximize_popup != nullptr) {
          maximize_popup->set_delete_after_restore(closeDesktopOnRestoreLastWindow.value);
        }
      }
      if (_values.is_int_value(popupDelay.name)) {
        popupDelay.value = _values.get_int_value(popupDelay.name);
        if (_enabled) {
          update_mousein_wait(popupDelay.value);
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
      maximize_popup->set_delete_after_restore(closeDesktopOnRestoreLastWindow.value);
      start_mouse_watcher(popupDelay.value, 100, on_mouse_in, on_mouse_out, maximize_popup->get_hwnd());
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
  void init_settings();
  void save_settings();

  // Settings used by this module
  struct CloseDesktopOnRestoreLastWindow {
    PCWSTR name = L"close_desktop_on_restore";
    BOOL value = TRUE;
    int resourceId = IDS_SETTING_DESCRIPTION_CLOSE_ON_RESTORE;
  } closeDesktopOnRestoreLastWindow;

  struct PopupDelay {
    PCWSTR name = L"popup_delay";
    int value = 400; // ms
    int resourceId = IDS_SETTING_DESCRIPTION_HOVER_DELAY;
  } popupDelay;
};

D2DWindowManagerPopup* MTNDPowertoy::maximize_popup = nullptr;
std::chrono::steady_clock::time_point shown_start_time;

RECT on_mouse_in(HWND hwnd, RECT buttons, POINT mouse_pos) {
  const int DEFAULT_DPI = 96;
  // The original svg image size is 35px wide and 13px high, that we
  // increase by 1.6 times
  const int svg_icon_width = (int)(35 * 1.6f);
  const int svg_icon_height = (int)(13 * 1.6f);
  // Add extra 8px that will be used for the padding
  int popup_width = svg_icon_width + 8;
  int popup_height = svg_icon_height + 8;

  HMONITOR monitor_handle = MonitorFromPoint(mouse_pos, MONITOR_DEFAULTTONEAREST);
  MonitorInfo monitor_info = get_monitor_info(monitor_handle);
  DPIAware::Convert(monitor_handle, popup_width, popup_height);

  LONG midPoint = (buttons.left + buttons.right) / 2;
  RECT result;
  result.left = midPoint - popup_width / 2;
  result.top = buttons.bottom;
  result.right = midPoint + popup_width / 2;
  result.bottom = buttons.bottom + popup_height;
  // Make sure resulting rect is inside monitor.
  result = keep_rect_inside_rect(result, monitor_info.rect);
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
        closeDesktopOnRestoreLastWindow.name,
        closeDesktopOnRestoreLastWindow.value
      )
    );
    values.add_property_value(
      PowerToysSettings::IntSettingsValue(
        popupDelay.name,
        popupDelay.value
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
    if (settings.is_bool_value(closeDesktopOnRestoreLastWindow.name)) {
      closeDesktopOnRestoreLastWindow.value = settings.get_bool_value(closeDesktopOnRestoreLastWindow.name);
    }
    if (settings.is_int_value(popupDelay.name)) {
      popupDelay.value = settings.get_int_value(popupDelay.name);
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
