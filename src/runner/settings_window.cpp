#include "pch.h"
#include <WinSafer.h>
#include <Sddl.h>
#include <sstream>
#include <accctrl.h>
#include <aclapi.h>
#include <cpprest/json.h>
#include "powertoy_module.h"
#include <common/two_way_pipe_message_ipc.h>
#include "tray_icon.h"
#include "general_settings.h"

#define BUFSIZE 1024

using namespace web;

TwoWayPipeMessageIPC* current_settings_ipc = NULL;

json::value get_power_toys_settings() {
  json::value result = json::value::object();
  for (auto&[name, powertoy] : modules()) {
    try {
      json::value powertoys_config = json::value::parse(powertoy.get_config());
      result.as_object()[name] = powertoys_config;
    }
    catch (json::json_exception &e) {
      //Malformed JSON.
    }
  }
  return result;
}

json::value get_all_settings() {
  json::value result = json::value::object();
  result.as_object()[L"general"] = get_general_settings();
  result.as_object()[L"powertoys"] = get_power_toys_settings();
  return result;
}

void send_json_config_to_module(const std::wstring& module_key, const std::wstring& settings) {
  if (modules().find(module_key) != modules().end()) {
    modules().at(module_key).set_config(settings);
  }
}

void dispatch_json_config_to_modules(const json::value& powertoys_configs) {
  for (auto powertoy_element : powertoys_configs.as_object()) {
    std::wstringstream ws;
    ws << powertoy_element.second;
    send_json_config_to_module(powertoy_element.first, ws.str());
  }
};

void dispatch_received_json(const std::wstring &json_to_parse) {
  json::value j = json::value::parse(json_to_parse);
  for(auto base_element : j.as_object()) {
    if (base_element.first == L"general") {
      apply_general_settings(base_element.second);
      std::wstringstream ws;
      ws << get_all_settings();
      if (current_settings_ipc != NULL) {
        current_settings_ipc->send(ws.str());
      }
    } else if (base_element.first == L"powertoys") {
      dispatch_json_config_to_modules(base_element.second);
      std::wstringstream ws;
      ws << get_all_settings();
      if (current_settings_ipc != NULL) {
        current_settings_ipc->send(ws.str());
      }
    } else if (base_element.first == L"refresh") {
      std::wstringstream ws;
      ws << get_all_settings();
      if (current_settings_ipc != NULL) {
        current_settings_ipc->send(ws.str());
      }
    }
  }
  return;
}

void dispatch_received_json_callback(PVOID data) {
  std::wstring* msg = (std::wstring*)data;
  dispatch_received_json(*msg);
  delete msg;
}

void receive_json_send_to_main_thread(const std::wstring &msg) {
  std::wstring* copy = new std::wstring(msg);
  dispatch_run_on_main_ui_thread(dispatch_received_json_callback, copy);
}

HANDLE create_medium_integrity_token() {
  HANDLE restricted_token_handle;
  SAFER_LEVEL_HANDLE level_handle = NULL;
  DWORD sid_size = SECURITY_MAX_SID_SIZE;
  BYTE medium_sid[SECURITY_MAX_SID_SIZE];
  if (!SaferCreateLevel(SAFER_SCOPEID_USER, SAFER_LEVELID_NORMALUSER, SAFER_LEVEL_OPEN, &level_handle, NULL)) {
    show_last_error_message(TEXT("Can't create a safer level token"), GetLastError());
    return NULL;
  }
  if (!SaferComputeTokenFromLevel(level_handle, NULL, &restricted_token_handle, 0, NULL)) {
    SaferCloseLevel(level_handle);
    show_last_error_message(TEXT("Can't create the restricted token"), GetLastError());
    return NULL;
  }
  SaferCloseLevel(level_handle);

  if (!CreateWellKnownSid(WinMediumLabelSid, nullptr, medium_sid, &sid_size)) {
    CloseHandle(restricted_token_handle);
    show_last_error_message(TEXT("Can't create a SID for medium integrity"), GetLastError());
    return NULL;
  }

  TOKEN_MANDATORY_LABEL integrity_level = { 0 };
  integrity_level.Label.Attributes = SE_GROUP_INTEGRITY;
  integrity_level.Label.Sid = reinterpret_cast<PSID>(medium_sid);

  if (!SetTokenInformation(restricted_token_handle, TokenIntegrityLevel, &integrity_level, sizeof(integrity_level))) {
    CloseHandle(restricted_token_handle);
    show_last_error_message(TEXT("Can't set the token integrity level to medium"), GetLastError());
    return NULL;
  }

  return restricted_token_handle;
}

bool block_settings_window_start = false;

void run_settings_window() {
  block_settings_window_start = true;
  STARTUPINFO startup_info = { sizeof(startup_info) };
  PROCESS_INFORMATION process_info = { 0 };
  TCHAR executable_path[MAX_PATH];
  GetModuleFileName(NULL, executable_path, MAX_PATH);
  PathRemoveFileSpec(executable_path);
  wcscat_s(executable_path, TEXT("\\settings.exe"));
  HANDLE restricted_token;
  TCHAR executable_args[MAX_PATH * 3];
  // Generate unique names for the pipes, if getting a UUID is possible
  std::wstring powertoys_pipe_name(TEXT("\\\\.\\pipe\\powertoys_runner_"));
  std::wstring settings_pipe_name(TEXT("\\\\.\\pipe\\powertoys_settings_"));
  UUID temp_uuid;
  UuidCreate(&temp_uuid);
  wchar_t* uuid_chars;
  UuidToString(&temp_uuid, (RPC_WSTR*)&uuid_chars);
  if (uuid_chars != NULL) {
    powertoys_pipe_name += std::wstring(uuid_chars);
    settings_pipe_name += std::wstring(uuid_chars);
    RpcStringFree((RPC_WSTR*)&uuid_chars);
    uuid_chars = NULL;
  }
  wcscpy_s(executable_args, TEXT("\""));
  wcscat_s(executable_args, executable_path);
  wcscat_s(executable_args, TEXT("\""));
  wcscat_s(executable_args, TEXT(" "));
  wcscat_s(executable_args, powertoys_pipe_name.c_str());
  wcscat_s(executable_args, TEXT(" "));
  wcscat_s(executable_args, settings_pipe_name.c_str());

  // TODO: Check integrity level before creating the token.

  // Create a restricted token. The WebView created by the settings can't run as administrator.
  restricted_token = create_medium_integrity_token();

  if (!restricted_token) {
    // Couldn't get the restricted token to spawn the new process.
    block_settings_window_start = false;
    return;
  }

  if (!CreateProcessAsUser(restricted_token, executable_path, executable_args, NULL, NULL, TRUE, CREATE_SUSPENDED, NULL, NULL, &startup_info, &process_info)) {
    show_last_error_message(TEXT("Can't open Settings Window"), GetLastError());
    block_settings_window_start = false;
    return;
  }

  current_settings_ipc = new TwoWayPipeMessageIPC(powertoys_pipe_name, settings_pipe_name, receive_json_send_to_main_thread);
  current_settings_ipc->start(restricted_token);

  ResumeThread(process_info.hThread);
  CloseHandle(process_info.hThread);

  if (WaitForSingleObject(process_info.hProcess, INFINITE) != WAIT_OBJECT_0) {
    show_last_error_message(TEXT("Couldn't wait on the Settings Window to close."), GetLastError());
  }

  current_settings_ipc->end();
  delete current_settings_ipc;
  current_settings_ipc = NULL;

  CloseHandle(restricted_token);

  CloseHandle(process_info.hProcess);
  block_settings_window_start = false;
}

void open_settings_window() {
  if (block_settings_window_start) {
    MessageBox(NULL, L"There's a Settings Window already running. Close the first instance first.", L"Settings", MB_OK && MB_TOPMOST);
  } else {
    block_settings_window_start = true;
    std::thread(run_settings_window).detach();
  }
}
