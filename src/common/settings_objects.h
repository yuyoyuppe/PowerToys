#pragma once
#include <string>
#include <cpprest/json.h>

namespace PowerToysSettings {

  class Settings;
  class BasePropertySetting;
  class BaseSettingsValue;

  class Settings {
  public:
    Settings(
      const std::wstring& powertoy_name,
      const std::wstring& description
    );

    void add_property(PowerToysSettings::BasePropertySetting _setting);

    std::wstring get_name();
    web::json::value get_json();
    std::wstring to_string();

    // Helper functions to get allocated strings out of the DLLs
    wchar_t* to_allocated_cstring();
    static void free_allocated_cstring(const wchar_t* allocated_cstring);

    // Add additional general information to the PowerToy settings.
    void set_icon_key(const std::wstring& icon_key);

  private:
    web::json::value _internal_json;
    std::wstring _name;
    int curr_priority = 0; // For keeping order when adding elements.
  };

  class BasePropertySetting {
  public:
    BasePropertySetting(
      const std::wstring& name,
      const std::wstring& display_name,
      const std::wstring& editor_type,
      web::json::value value
    );

    std::wstring get_name();
    web::json::value get_json();
  protected:
    web::json::value _internal_json;
    std::wstring _name;
  };

  class BoolTogglePropertySetting : public BasePropertySetting {
  public:
    BoolTogglePropertySetting(
      const std::wstring& name,
      const std::wstring& display_name,
      bool value
    ) : BasePropertySetting(
      name,
      display_name,
      L"bool_toggle",
      web::json::value::boolean(value)
    ) {
    }
  };

  class IntSpinnerPropertySetting : public BasePropertySetting {
  public:
    IntSpinnerPropertySetting(
      const std::wstring& name,
      const std::wstring& display_name,
      int value
    ) : BasePropertySetting(
      name,
      display_name,
      L"int_spinner",
      web::json::value::number(value)
    ) {
    }
    void setSpinnerMin(int min_value);
    void setSpinnerMax(int max_value);
    void setSpinnerStep(int step);
  };

  class StringTextPropertySetting : public BasePropertySetting {
  public:
    StringTextPropertySetting(
      const std::wstring& name,
      const std::wstring& display_name,
      const std::wstring& value
    ) : BasePropertySetting(
      name,
      display_name,
      L"string_text",
      web::json::value::string(value)
    ) {
    }
  };

  class ColorPickerPropertySetting : public BasePropertySetting {
  public:
    ColorPickerPropertySetting(
      const std::wstring& name,
      const std::wstring& display_name,
      const std::wstring& value
    ) : BasePropertySetting(
      name,
      display_name,
      L"color_picker",
      web::json::value::string(value)
    ) {
    }
  };

  class CustomActionPropertySetting : public BasePropertySetting {
  public:
    CustomActionPropertySetting(
      const std::wstring& name,
      const std::wstring& display_name,
      const std::wstring& value,
      const std::wstring& button_text
    ) : BasePropertySetting (
      name,
      display_name,
      L"custom_action",
      web::json::value::string(value)
    ) {
      _internal_json.as_object()[L"button_text"] = web::json::value::string(button_text);
    }
  };

  class PowerToyValues {
  public:
    PowerToyValues(const std::wstring& powertoy_name);
    static PowerToyValues from_json_string(const std::wstring& json);
    static PowerToyValues load_from_settings_file(const std::wstring& powertoy_name);

    void add_property_value(PowerToysSettings::BaseSettingsValue _value);

    std::wstring get_name();
    web::json::value get_json();
    std::wstring to_string();

    // Check property value type
    bool is_bool_value(const std::wstring& property_name);
    bool is_int_value(const std::wstring& property_name);
    bool is_string_value(const std::wstring& property_name);

    // Get property value
    bool get_bool_value(const std::wstring& property_name);
    int get_int_value(const std::wstring& property_name);
    std::wstring get_string_value(const std::wstring& property_name);

    void save_to_settings_file();

  private:
    web::json::value _internal_json;
    std::wstring _name;
    PowerToyValues() {}
  };

  class BaseSettingsValue {
  public:
    BaseSettingsValue(
      const std::wstring& name,
      web::json::value value
    );
    std::wstring get_name();
    web::json::value get_json();
  protected:
    web::json::value _internal_json;
    std::wstring _name;
  };

  class BoolSettingsValue : public BaseSettingsValue {
  public:
    BoolSettingsValue(
      const std::wstring& name,
      bool value
    ) : BaseSettingsValue(
      name,
      web::json::value::boolean(value)
    ) {
    }
  };

  class IntSettingsValue : public BaseSettingsValue {
  public:
    IntSettingsValue(
      const std::wstring& name,
      int value
    ) : BaseSettingsValue(
      name,
      web::json::value::number(value)
    ) {
    }
  };

  class StringSettingsValue : public BaseSettingsValue {
  public:
    StringSettingsValue(
      const std::wstring& name,
      std::wstring value
    ) : BaseSettingsValue(
      name,
      web::json::value::string(value)
    ) {
    }
  };

  class CustomActionObject {
  public:
    static CustomActionObject from_json_string(const std::wstring& json);
    std::wstring get_name();
    std::wstring get_value();
    web::json::value get_json();
  protected:
    CustomActionObject(
      web::json::value value
    );
    web::json::value _internal_json;
  };

}
