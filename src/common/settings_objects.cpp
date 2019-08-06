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

  std::wstring Settings::to_string() {
    return _internal_json.serialize();
  }

  wchar_t* Settings::to_allocated_cstring() {
    std::wstring w_str = to_string();
    size_t result_len = w_str.length() + 1;
    wchar_t *result = new wchar_t[result_len];
    wcscpy_s(result, result_len, w_str.c_str());
    return result;
  }

  void Settings::free_allocated_cstring(const wchar_t * allocated_cstring) {
    delete[] allocated_cstring;
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
    _internal_json.as_object()[L"version"] = web::json::value::string(L"1.0");
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

  void PowerToyValues::add_property_value(BaseSettingsValue _value) {
    web::json::value property_json = _value.get_json();
    _internal_json.as_object()[L"properties"].as_object()[_value.get_name()] = property_json;
  }

  std::wstring PowerToyValues::get_name() {
    return _name;
  }

  web::json::value PowerToyValues::get_json() {
    return _internal_json;
  }

  std::wstring PowerToyValues::to_string() {
    return _internal_json.serialize();
  }

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
    PTSettingsHelper::save_module_settings(_name, _internal_json);
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