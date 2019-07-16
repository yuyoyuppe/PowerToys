#pragma once
#include <string>
#include <Shlobj.h>
#include <cpprest/json.h>
namespace PowerToysSettings {
  std::wstring get_global_powertoys_save_folder_location();
  std::wstring get_powertoy_save_folder_location(std::wstring& powertoy_name);
  std::wstring get_powertoy_save_file_location(std::wstring& powertoy_name);
  std::wstring get_powertoys_general_save_file_location();
  void save_powertoy_settings_json(std::wstring& powertoy_name, web::json::value& settings);
  web::json::value load_powertoy_settings_json(std::wstring& powertoy_name);
  void save_general_settings_json(web::json::value& settings);
  web::json::value load_general_settings_json();


}
