#include "pch.h"
#include "settings_helpers.h"
#include <filesystem>
#include <fstream>

std::wstring PowerToysSettings::get_global_powertoys_save_folder_location() {
  PWSTR local_app_path;
  std::wstring result(L"");
  if (SUCCEEDED(SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, NULL, &local_app_path))) {
    result=std::wstring(local_app_path);
    CoTaskMemFree(local_app_path);
  }
  result += L"\\Microsoft\\PowerToys";
  std::filesystem::path save_path(result);
  if (!std::filesystem::exists(save_path)) {
    std::filesystem::create_directories(save_path);
  }
  return result;
}

std::wstring PowerToysSettings::get_powertoy_save_folder_location(std::wstring& powertoy_name) {
  std::wstring result = get_global_powertoys_save_folder_location();
  result += L"\\";
  result += powertoy_name;
  std::filesystem::path save_path(result);
  if (!std::filesystem::exists(save_path)) {
    std::filesystem::create_directories(save_path);
  } 
  return result;
}

std::wstring PowerToysSettings::get_powertoy_save_file_location(std::wstring& powertoy_name) {
  std::wstring result = get_powertoy_save_folder_location(powertoy_name);
  result += L"\\settings.json";
  return result;
}

std::wstring PowerToysSettings::get_powertoys_general_save_file_location() {
  std::wstring result = get_global_powertoys_save_folder_location();
  result += L"\\settings.json";
  return result;
}

void PowerToysSettings::save_powertoy_settings_json(std::wstring& powertoy_name, web::json::value& settings) {
  std::wstring save_file_location = get_powertoy_save_file_location(powertoy_name);
  std::ofstream save_file(save_file_location, std::ios::binary);
  settings.serialize(save_file);
  save_file.close();
}

web::json::value PowerToysSettings::load_powertoy_settings_json(std::wstring& powertoy_name) {
  std::wstring save_file_location = get_powertoy_save_file_location(powertoy_name);
  std::ifstream save_file(save_file_location, std::ios::binary);
  web::json::value result = web::json::value::parse(save_file);
  save_file.close();
  return result;
}

void PowerToysSettings::save_general_settings_json(web::json::value& settings) {
  std::wstring save_file_location = get_powertoys_general_save_file_location();
  std::ofstream save_file(save_file_location, std::ios::binary);
  settings.serialize(save_file);
  save_file.close();
}

web::json::value PowerToysSettings::load_general_settings_json() {
  std::wstring save_file_location = get_powertoys_general_save_file_location();
  std::ifstream save_file(save_file_location, std::ios::binary);
  web::json::value result = web::json::value::parse(save_file);
  save_file.close();
  return result;
}
