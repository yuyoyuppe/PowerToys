#include "pch.h"
#include "settings_objects.h"
#include "settings_helpers.h"

namespace PowerToysSettings {

  Settings::Settings(const std::wstring& powertoy_name, const std::wstring& description) {
    _internal_json = web::json::value::object();
    _internal_json.as_object()[L"version"] = web::json::value::string(L"1.0");
    _internal_json.as_object()[L"name"] = web::json::value::string(powertoy_name);
    _internal_json.as_object()[L"description"] = web::json::value::string(description);
    _internal_json.as_object()[L"properties"] = web::json::value::object();
  }

  void Settings::add_property(BasePropertySetting _setting) {
    web::json::value property_json = _setting.get_json();
    property_json.as_object()[L"order"] = web::json::value::number(++(this->curr_priority));
    _internal_json.as_object()[L"properties"].as_object()[_setting.get_name()] = property_json;
  }

  bool Settings::serialize_to_buffer(wchar_t* buffer, int *buffer_size) {
    std::wstring result = _internal_json.serialize();
    int result_len = (int)result.length();

    if (buffer == nullptr || *buffer_size < result_len) {
      *buffer_size = result_len + 1;
      return false;
    } else {
      wcscpy_s(buffer, *buffer_size, result.c_str());
      return true;
    }
  }

  void Settings::set_icon_key(const std::wstring & icon_key) {
    _internal_json.as_object()[L"icon_key"] = web::json::value::string(icon_key);
  }

  void Settings::set_overview_link(const std::wstring & overview_link) {
    _internal_json.as_object()[L"overview_link"] = web::json::value::string(overview_link);
  }

  void Settings::set_video_link(const std::wstring & video_link) {
    _internal_json.as_object()[L"video_link"] = web::json::value::string(video_link);
  }

  BasePropertySetting::BasePropertySetting(const std::wstring& name, const std::wstring& display_name, const std::wstring& editor_type, web::json::value value) {
    _name = name;
    _internal_json = web::json::value::object();
    _internal_json.as_object()[L"display_name"] = web::json::value::string(display_name);
    _internal_json.as_object()[L"editor_type"] = web::json::value::string(editor_type);
    _internal_json.as_object()[L"value"] = value;
  }

  std::wstring BasePropertySetting::get_name() {
    return _name;
  }

  web::json::value BasePropertySetting::get_json() {
    return _internal_json;
  }

  PowerToyValues::PowerToyValues(const std::wstring& powertoy_name) {
    _name = powertoy_name;
    _internal_json = web::json::value::object();
    set_version();
    _internal_json.as_object()[L"name"] = web::json::value::string(powertoy_name);
    _internal_json.as_object()[L"properties"] = web::json::value::object();
  }

  PowerToyValues PowerToyValues::from_json_string(const std::wstring& json) {
    PowerToyValues result = PowerToyValues();
    result._internal_json = web::json::value::parse(json);
    result._name = result._internal_json.as_object()[L"name"].as_string();
    return result;
  }

  PowerToyValues PowerToyValues::load_from_settings_file(const std::wstring & powertoy_name) {
    PowerToyValues result = PowerToyValues();
    result._internal_json = PTSettingsHelper::load_module_settings(powertoy_name);
    result._name = powertoy_name;
    return result;
  }

  template <typename T>
  web::json::value add_property_generic(const std::wstring& name, T value) {
    std::vector<std::pair<std::wstring, web::json::value>> vector = { std::make_pair(L"value", web::json::value(value)) };
    return web::json::value::object(vector);
  }

  template <>
  void PowerToyValues::add_property(const std::wstring& name, bool value) {
    _internal_json.as_object()[L"properties"].as_object()[name] = add_property_generic(name, value);
  };
  
  template <>
  void PowerToyValues::add_property(const std::wstring& name, int value) {
    _internal_json.as_object()[L"properties"].as_object()[name] = add_property_generic(name, value);
  };

  template <>
  void PowerToyValues::add_property(const std::wstring& name, std::wstring value) {
    _internal_json.as_object()[L"properties"].as_object()[name] = add_property_generic(name, value);
  };

  bool PowerToyValues::is_bool_value(const std::wstring& property_name) {
    return _internal_json.is_object() &&
      _internal_json.has_object_field(L"properties") &&
      _internal_json[L"properties"].has_object_field(property_name) &&
      _internal_json[L"properties"][property_name].has_boolean_field(L"value");
  }

  bool PowerToyValues::is_int_value(const std::wstring& property_name) {
    return _internal_json.is_object() &&
      _internal_json.has_object_field(L"properties") &&
      _internal_json[L"properties"].has_object_field(property_name) &&
      _internal_json[L"properties"][property_name].has_integer_field(L"value");
  }

  bool PowerToyValues::is_string_value(const std::wstring& property_name) {
    return _internal_json.is_object() &&
      _internal_json.has_object_field(L"properties") &&
      _internal_json[L"properties"].has_object_field(property_name) &&
      _internal_json[L"properties"][property_name].has_string_field(L"value");
  }

  bool PowerToyValues::get_bool_value(const std::wstring& property_name) {
    return _internal_json[L"properties"][property_name][L"value"].as_bool();
  }

  int PowerToyValues::get_int_value(const std::wstring& property_name) {
    return _internal_json[L"properties"][property_name][L"value"].as_integer();
  }

  std::wstring PowerToyValues::get_string_value(const std::wstring& property_name) {
    return _internal_json[L"properties"][property_name][L"value"].as_string();
  }

  void PowerToyValues::save_to_settings_file() {
    set_version();
    PTSettingsHelper::save_module_settings(_name, _internal_json);
  }

  void PowerToyValues::set_version() {
    _internal_json.as_object()[L"version"] = web::json::value::string(m_version);
  }

  BaseSettingsValue::BaseSettingsValue(const std::wstring & name, web::json::value value) {
    _name = name;
    _internal_json = web::json::value::object();
    _internal_json.as_object()[L"value"] = value;
  }

  std::wstring BaseSettingsValue::get_name() {
    return _name;
  }

  web::json::value BaseSettingsValue::get_json() {
    return _internal_json;
  }

  void IntSpinnerPropertySetting::setSpinnerMin(int min_value) {
    _internal_json.as_object()[L"min"] = web::json::value::number(min_value);
  }

  void IntSpinnerPropertySetting::setSpinnerMax(int max_value) {
    _internal_json.as_object()[L"max"] = web::json::value::number(max_value);
  }

  void IntSpinnerPropertySetting::setSpinnerStep(int step) {
    _internal_json.as_object()[L"step"] = web::json::value::number(step);
  }

  CustomActionObject::CustomActionObject(web::json::value action_json) {
    _internal_json = action_json;
  }

  CustomActionObject CustomActionObject::from_json_string(const std::wstring & json) {
    web::json::value parsed_json = web::json::value::parse(json);
    return CustomActionObject(parsed_json);
  }

  std::wstring CustomActionObject::get_name() {
    return _internal_json[L"action_name"].as_string();
  }

  std::wstring CustomActionObject::get_value() {
    return _internal_json[L"value"].as_string();
  }

  web::json::value CustomActionObject::get_json() {
    return _internal_json;
  }

}