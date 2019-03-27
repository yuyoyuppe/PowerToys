#include <windows.h>
#include <ShellScalingApi.h>
#pragma comment(lib, "shcore.lib")
#include "functionalities.h"
#include "utils.h"

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
  // We will handle scaling ourselfs.
  SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE);
 
  start_winkey_handler();
  start_mouse_handler();
  return run_message_loop();
}
