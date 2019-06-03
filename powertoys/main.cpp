#include "pch.h"
#include "functionalities.h"
#include <ShellScalingApi.h>
#include "d2d_window.h"
#pragma comment(lib, "shcore.lib")
#pragma comment(lib, "windowsapp")

#if _DEBUG && _WIN64
#include "unhandled_exception_handler.h"
#endif

void chdir_current_executable() {
  // Change current directory to the path of the executable.
  CCHAR executable_path[MAX_PATH];
  GetModuleFileName(NULL, executable_path, MAX_PATH);
  PathRemoveFileSpec(executable_path);
  if(!SetCurrentDirectory(executable_path)) {
    show_last_error_message((LPTSTR)"Change Directory to Executable Path", GetLastError());
  }
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
  #if _DEBUG && _WIN64
  //Global error handlers to diagnose errors.
  //We prefer this not not show any longer until there's a bug to diagnose.
  //init_global_error_handlers();
  #endif
  winrt::init_apartment();
  winrt::check_hresult(SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE));
  // We will handle scaling ourselfs.
  HHOOK handle = NULL;
  try {
    chdir_current_executable();
    start_tray_icon();
    start_winkey_handler();
    start_mouse_handler();
    int result = run_message_loop();
  return result;
  } catch (std::runtime_error err) {
    MessageBox(NULL, err.what(), "Error", MB_OK | MB_ICONERROR);
    return -1;
  }
}
