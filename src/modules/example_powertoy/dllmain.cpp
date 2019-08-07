#include "pch.h"
#include <interface/powertoy_module_interface.h>
#include <interface/lowlevel_keyboard_event_data.h>
#include <interface/win_hook_event_data.h>
#include "trace.h"
#include <common/settings_objects.h>

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
  // PowerToys properties can be saved here.
  bool test_bool_prop = false;
  int test_int_prop = 3;
  std::wstring test_string_prop = L"a string";
  std::wstring test_color_prop = L"#1212FF";
  int test_custom_action_num_calls = 0;
  
  // Load and save Settings from/to local app data.
  void init_settings();
  void save_settings();
  
  // If the PowerToy is enabled.
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
    static const wchar_t* events[] = { ll_keyboard,
                                       win_hook_event,
                                       nullptr };
    return events;
  }
  // Return JSON with the configuration options.
  virtual bool get_config(wchar_t* buffer, int *buffer_size) override {
    // Create a Settings object.
    PowerToysSettings::Settings settings(
      get_name(),
      L"Serves as an example powertoy, with example settings."
    );

    // Add an overview link to show in the Settings.
    settings.set_overview_link(L"https://github.com/microsoft/PowerToys");

    // Add a video link to show in the Settings.
    settings.set_video_link(L"https://www.youtube.com/watch?v=d3LHo2yXKoY&t=21462");

    // Add a bool property with a toggle editor.
    settings.add_property(
      PowerToysSettings::BoolTogglePropertySetting(
        L"test bool_toggle", // property name
        L"This is what a BoolTogglePropertySetting looks like", // property display text
        test_bool_prop // property value
      )
    );

    // Add an integer property with a spinner editor.
    settings.add_property(
      PowerToysSettings::IntSpinnerPropertySetting(
        L"test int_spinner", // property name
        L"This is what a IntSpinnerPropertySetting looks like", // property display text
        test_int_prop // property value
      )
    );

    // Add a string property with a textbox editor.
    settings.add_property(
      PowerToysSettings::StringTextPropertySetting(
        L"test string_text", // property name
        L"This is what a StringTextPropertySetting looks like", // property display text
        test_string_prop // property value
      )
    );

    // Add a string property with a color picker editor.
    settings.add_property(
      PowerToysSettings::ColorPickerPropertySetting(
        L"test color_picker", // property name
        L"This is what a ColorPickerPropertySetting looks like", // property display text
        test_color_prop // property value
      )
    );

    // Add a custom action property. When using this settings type, "call_custom_action" should be overriden as well.
    settings.add_property(
      PowerToysSettings::CustomActionPropertySetting(
        L"test custom_action", // action name
        L"This is what a CustomActionPropertySetting looks like", // label above the field
        L"Press the button to call a custom action in the Example PowerToy", // display values / extended info
        L"Call a custom action!" // button text
        )
    );

    return settings.serialize_to_buffer(buffer, buffer_size);
    }

  // Signal from the settings screen to call a custom action.
  // This can be used to spawn more complex editors.
  virtual void call_custom_action(const wchar_t* action) override {
    try {
      // Parse the action values, including name.
      PowerToysSettings::CustomActionObject action_object =
        PowerToysSettings::CustomActionObject::from_json_string(action);

      if (action_object.get_name() == L"test custom_action") {

        // Custom action code to increase and show a counter.
        ++this->test_custom_action_num_calls;
        std::wstring msg(L"I have been called ");
        msg += std::to_wstring(this->test_custom_action_num_calls);
        msg += L" time(s).";
        MessageBox(NULL, msg.c_str(), L"Custom action call.", MB_OK | MB_TOPMOST);
      }
    }
    catch (std::exception ex) {
      // Improper JSON.
    }
  }

  // Passes JSON with the configuration settings for the powertoy.
  virtual void set_config(const wchar_t* config) override { 
    try {
      // Parse the PowerToysValues object from the received json string.
      PowerToysSettings::PowerToyValues _values =
        PowerToysSettings::PowerToyValues::from_json_string(config);

      // Update the bool property.
      if (_values.is_bool_value(L"test bool_toggle")) {
        test_bool_prop = _values.get_bool_value(L"test bool_toggle");
      }

      // Update the int property.
      if (_values.is_int_value(L"test int_spinner")) {
        test_int_prop = _values.get_int_value(L"test int_spinner");
      }

      // Update the string property.
      if (_values.is_string_value(L"test string_text")) {
        test_string_prop = _values.get_string_value(L"test string_text");
      }

      // Update the color property.
      if (_values.is_string_value(L"test color_picker")) {
        test_color_prop = _values.get_string_value(L"test color_picker");
      }

      // If you don't need to do any custom processing of the module settings,
      // to persists the values as they are simply call:

      //_values.save_to_settings_file();
      
      // Otherwise call a custom function to process the settings, recreate
      // a PowerToysSettings::PowerToyValues and then save to disk
      save_settings();
    }
    catch (std::exception ex) {
      // Improper JSON.
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
    if (wcscmp(name, ll_keyboard) == 0) {
      auto& event = *(reinterpret_cast<LowlevelKeyboardEvent*>(data));
      // Return 1 if the keypress is to be suppressed (not forwarded to Windows),
      // otherwise return 0.
      return 0;
    } else if (wcscmp(name, win_hook_event) == 0) {
      auto& event = *(reinterpret_cast<WinHookEvent*>(data));
      // Return value is ignored
      return 0;
    }
    return 0;
  }

  // Constructor
  ExamplePowertoy();

  // Destroy the powertoy and free memory
  virtual void destroy() override {
    delete this;
  }
};

// Constructor
ExamplePowertoy::ExamplePowertoy() {
  init_settings();
}

// Load the settings file.
void ExamplePowertoy::init_settings() {
  try {
    // Load and parse the settings file for this PowerToy.
    PowerToysSettings::PowerToyValues settings =
      PowerToysSettings::PowerToyValues::load_from_settings_file(get_name());

    // Load the bool property.
    if (settings.is_bool_value(L"test bool_toggle")) {
      test_bool_prop = settings.get_bool_value(L"test bool_toggle");
    }

    // Load the int property.
    if (settings.is_int_value(L"test int_spinner")) {
      test_int_prop = settings.get_int_value(L"test int_spinner");
    }

    // Load the string property.
    if (settings.is_string_value(L"test string_text")) {
      test_string_prop = settings.get_string_value(L"test string_text");
    }

    // Load the color property.
    if (settings.is_string_value(L"test color_picker")) {
      test_color_prop = settings.get_string_value(L"test color_picker");
    }
  }
  catch (std::exception ex) {
    // Error while loading from the settings file. Just let default values stay as they are.
  }
}

void ExamplePowertoy::save_settings() {
  try {
    // Create a PowerToyValues object for this PowerToy
    PowerToysSettings::PowerToyValues values(
      get_name()
    );

    // Save the bool property.
    values.add_property(
      L"test bool_toggle", // property name
      test_bool_prop // property value
    );

    // Save the int property.
    values.add_property(
      L"test int_spinner", // property name
      test_int_prop // property value
    );

    // Save the string property.
    values.add_property(
      L"test string_text", // property name
      test_string_prop // property value
    );

    // Save the color property.
    values.add_property(
      L"test color_picker", // property name
      test_color_prop // property value
    );

    // Save the PowerToyValues JSON to the power toy settings file.
    values.save_to_settings_file();

  }
  catch (std::exception ex) {
    //Couldn't save the settings.
  }
}

extern "C" __declspec(dllexport) PowertoyModuleIface*  __cdecl powertoy_create() {
  return new ExamplePowertoy();
}


