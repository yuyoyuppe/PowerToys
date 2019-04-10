#include "pch.h"
#include "functionalities.h"
#include <ShellScalingApi.h>
#pragma comment(lib, "shcore.lib")
#pragma comment(lib, "windowsapp")

#include "d2d_window.h"
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
  winrt::init_apartment();
  
  try {
    // We will handle scaling ourselfs.
    winrt::check_hresult(SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE));

    D2DWindow test;
    /*start_winkey_handler();
    start_mouse_handler();*/
    return run_message_loop();
  } catch (std::runtime_error err) {
    MessageBox(NULL, err.what(), "Error", MB_OK | MB_ICONERROR);
    return -1;
  }
}
