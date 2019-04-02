#include <windows.h>
#include <ShellScalingApi.h>
#pragma comment(lib, "shcore.lib")
#include <stdexcept>
#include "functionalities.h"
#include "utils.h"

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
  try {
    // We will handle scaling ourselfs.
    SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE);

    start_winkey_handler();
    //start_mouse_handler();
    return run_message_loop();
  } catch (std::runtime_error err) {
    MessageBox(NULL, err.what(), "Error", MB_OK | MB_ICONERROR);
    return -1;
  }
}
