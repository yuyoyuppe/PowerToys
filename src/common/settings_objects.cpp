#include "pch.h"
#include "settings_objects.h"
#include "settings_helpers.h"

PowerToysSettings::Settings::Settings(const std::wstring& powertoy_name, const std::wstring& description) {
  _name = powertoy_name;
  _internal_json = web::json::value::object();
  _internal_json.as_object()[L"version"] = web::json::value::string(L"1.0");
  _internal_json.as_object()[L"name"] = web::json::value::string(powertoy_name);
  _internal_json.as_object()[L"description"] = web::json::value::string(description);
  _internal_json.as_object()[L"properties"] = web::json::value::object();
}

void PowerToysSettings::Settings::add_property(PowerToysSettings::BasePropertySetting _setting) {
  web::json::value property_json = _setting.get_json();
  property_json.as_object()[L"order"] = web::json::value::number(++(this->curr_priority));
  _internal_json.as_object()[L"properties"].as_object()[_setting.get_name()] = property_json;
}

std::wstring PowerToysSettings::Settings::get_name() {
  return _name;
}

web::json::value PowerToysSettings::Settings::get_json() {
  return _internal_json;
}

std::wstring PowerToysSettings::Settings::to_string() {
  return _internal_json.serialize();
}

wchar_t * PowerToysSettings::Settings::to_allocated_cstring() {
  std::wstringstream ss;
  ss << to_string();
  std::wstring w_str = ss.str();
  const wchar_t* result_wstr = w_str.c_str();
  size_t result_len = wcslen(result_wstr) + 1;
  wchar_t *result = new wchar_t[result_len];
  wcscpy_s(result, result_len, result_wstr);
  return result;
}

void PowerToysSettings::Settings::free_allocated_cstring(const wchar_t * allocated_cstring) {
  delete[] allocated_cstring;
}

void PowerToysSettings::Settings::set_icon_key(const std::wstring & icon_key) {
  _internal_json.as_object()[L"icon_key"] = web::json::value::string(icon_key);
}

PowerToysSettings::BasePropertySetting::BasePropertySetting(const std::wstring& name, const std::wstring& display_name, const std::wstring& editor_type, web::json::value value) {
  _name = name;
  _internal_json = web::json::value::object();
  _internal_json.as_object()[L"display_name"] = web::json::value::string(display_name);
  _internal_json.as_object()[L"editor_type"] = web::json::value::string(editor_type);
  _internal_json.as_object()[L"value"] = value;
}

std::wstring PowerToysSettings::BasePropertySetting::get_name() {
  return _name;
}

web::json::value PowerToysSettings::BasePropertySetting::get_json() {
  return _internal_json;
}

PowerToysSettings::PowerToyValues::PowerToyValues(const std::wstring& powertoy_name) {
  _name = powertoy_name;
  _internal_json = web::json::value::object();
  _internal_json.as_object()[L"version"] = web::json::value::string(L"1.0");
  _internal_json.as_object()[L"name"] = web::json::value::string(powertoy_name);
  _internal_json.as_object()[L"properties"] = web::json::value::object();
}

PowerToysSettings::PowerToyValues PowerToysSettings::PowerToyValues::from_json_string(const std::wstring& json) {
  PowerToysSettings::PowerToyValues result = PowerToysSettings::PowerToyValues();
  result._internal_json = web::json::value::parse(json);
  result._name = result._internal_json.as_object()[L"name"].as_string();
  return result;
}

PowerToysSettings::PowerToyValues PowerToysSettings::PowerToyValues::load_from_settings_file(const std::wstring & powertoy_name) {
  PowerToysSettings::PowerToyValues result = PowerToysSettings::PowerToyValues();
  result._internal_json = PowerToysSettings::load_powertoy_settings_json(powertoy_name);
  result._name = powertoy_name;
  return result;
}

void PowerToysSettings::PowerToyValues::add_property_value(PowerToysSettings::BaseSettingsValue _value) {
  web::json::value property_json = _value.get_json();
  _internal_json.as_object()[L"properties"].as_object()[_value.get_name()] = property_json;
}

std::wstring PowerToysSettings::PowerToyValues::get_name() {
  return _name;
}

web::json::value PowerToysSettings::PowerToyValues::get_json() {
  return _internal_json;
}

std::wstring PowerToysSettings::PowerToyValues::to_string() {
  return _internal_json.serialize();
}

bool PowerToysSettings::PowerToyValues::is_bool_value(const std::wstring& property_name) {
  return _internal_json.is_object() &&
    _internal_json.has_object_field(L"properties") &&
    _internal_json[L"properties"].has_object_field(property_name) &&
    _internal_json[L"properties"][property_name].has_boolean_field(L"value");
}

bool PowerToysSettings::PowerToyValues::is_int_value(const std::wstring& property_name) {
  return _internal_json.is_object() &&
    _internal_json.has_object_field(L"properties") &&
    _internal_json[L"properties"].has_object_field(property_name) &&
    _internal_json[L"properties"][property_name].has_integer_field(L"value");
}

bool PowerToysSettings::PowerToyValues::is_string_value(const std::wstring& property_name) {
  return _internal_json.is_object() &&
    _internal_json.has_object_field(L"properties") &&
    _internal_json[L"properties"].has_object_field(property_name) &&
    _internal_json[L"properties"][property_name].has_string_field(L"value");
}

bool PowerToysSettings::PowerToyValues::get_bool_value(const std::wstring& property_name) {
  return _internal_json[L"properties"][property_name][L"value"].as_bool();
}

int PowerToysSettings::PowerToyValues::get_int_value(const std::wstring& property_name) {
  return _internal_json[L"properties"][property_name][L"value"].as_integer();
}

std::wstring PowerToysSettings::PowerToyValues::get_string_value(const std::wstring& property_name) {
  return _internal_json[L"properties"][property_name][L"value"].as_string();
}

void PowerToysSettings::PowerToyValues::save_to_settings_file() {
  PowerToysSettings::save_powertoy_settings_json(_name, _internal_json);
}

PowerToysSettings::BaseSettingsValue::BaseSettingsValue(const std::wstring & name, web::json::value value) {
  _name = name;
  _internal_json = web::json::value::object();
  _internal_json.as_object()[L"value"] = value;
}

std::wstring PowerToysSettings::BaseSettingsValue::get_name() {
  return _name;
}

web::json::value PowerToysSettings::BaseSettingsValue::get_json() {
  return _internal_json;
}

void PowerToysSettings::IntSpinnerPropertySetting::setSpinnerMin(int min_value) {
  _internal_json.as_object()[L"min"] = web::json::value::number(min_value);
}

void PowerToysSettings::IntSpinnerPropertySetting::setSpinnerMax(int max_value) {
  _internal_json.as_object()[L"max"] = web::json::value::number(max_value);
}

void PowerToysSettings::IntSpinnerPropertySetting::setSpinnerStep(int step) {
  _internal_json.as_object()[L"step"] = web::json::value::number(step);
}

PowerToysSettings::CustomActionObject::CustomActionObject(web::json::value action_json) {
  _internal_json = action_json;
}

PowerToysSettings::CustomActionObject PowerToysSettings::CustomActionObject::from_json_string(const std::wstring & json) {
  web::json::value parsed_json = web::json::value::parse(json);
  return PowerToysSettings::CustomActionObject(parsed_json);
}

std::wstring PowerToysSettings::CustomActionObject::get_name() {
  return _internal_json[L"action_name"].as_string();
}

std::wstring PowerToysSettings::CustomActionObject::get_value() {
  return _internal_json[L"value"].as_string();
}

web::json::value PowerToysSettings::CustomActionObject::get_json() {
  return _internal_json;
}
