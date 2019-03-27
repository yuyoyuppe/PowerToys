#include <windows.h>
#include <ShellScalingApi.h>
#pragma comment(lib, "shcore.lib")

#include "functionalities.h"
#include "utils.h"
/*
void maximize_thread() {
  using namespace std::chrono_literals;
  while (true) {
    std::this_thread::sleep_for(100ms);
    POINT mouse_pos;
    GetCursorPos(&mouse_pos);
    auto mouse_window = WindowFromPoint(mouse_pos);
    if (mouse_window == NULL) {
      if (!start_pressed) {
        window->hide();
      }
      continue;
    }
    RECT window_rect;
    GetWindowRect(mouse_window, &window_rect);
    if (mouse_pos.x > window_rect.right - 200 &&
        mouse_pos.y < window_rect.top + 50) {
      window->show(window_rect.right - 220, window_rect.top + 50, 200, 30);
    } else if (!start_pressed) {
      window->hide();
    }
  }
}

std::thread(maximize_thread).detach();
*/

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
  // We will handle scaling ourselfs.
  SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE);
  start_winkey_handler();
  return run_message_loop();
}
