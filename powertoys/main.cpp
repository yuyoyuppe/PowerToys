#include "pch.h"
#include "functionalities.h"
#include <ShellScalingApi.h>
#include "d2d_window.h"
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
  winrt::init_apartment();
  winrt::check_hresult(SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE));
  try {
    start_winkey_handler();
    return run_message_loop();
  } catch (std::runtime_error err) {
    MessageBox(NULL, err.what(), "Error", MB_OK | MB_ICONERROR);
    return -1;
  }
}
