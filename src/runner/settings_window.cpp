#include "pch.h"
#include <WinSafer.h>
#include <Sddl.h>

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

void open_settings_window() {
  STARTUPINFO startup_info = { sizeof(startup_info) };
  PROCESS_INFORMATION process_info = { 0 };
  TCHAR executable_path[MAX_PATH];
  GetModuleFileName(NULL, executable_path, MAX_PATH);
  PathRemoveFileSpec(executable_path);
  wcscat_s(executable_path, TEXT("\\settings.exe"));
  HANDLE restricted_token;

  // TODO: Check integrity level before creating the token.

  // Create a restricted token. The WebView created by the settings can't run as administrator.
  restricted_token = create_medium_integrity_token();
  
  if (!restricted_token) {
    // Couldn't get the restricted token to spawn the new process.
    return;
  }

  if (CreateProcessAsUser(restricted_token, executable_path, executable_path, NULL, NULL, TRUE, 0, NULL, NULL, &startup_info, &process_info)) {
    CloseHandle(process_info.hProcess);
    CloseHandle(process_info.hThread);
  } else {
    show_last_error_message(TEXT("Can't open Settings Window"), GetLastError());
  }

  /* Cleanup */
  CloseHandle(restricted_token);
}
