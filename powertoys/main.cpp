#include "pch.h"
#include "functionalities.h"
#include <ShellScalingApi.h>
#include "d2d_window.h"
#pragma comment(lib, "shcore.lib")
#pragma comment(lib, "windowsapp")

#if _DEBUG && _WIN64
#include "unhandled_exception_handler.h"
#endif

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
  #if _DEBUG && _WIN64
  //Global error handlers to diagnose errors.
    InitGlobalErrorHandlers();
  #endif
  winrt::init_apartment();
  winrt::check_hresult(SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE));
  // We will handle scaling ourselfs.
  HHOOK handle = NULL;
  try {
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
