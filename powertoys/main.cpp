#include <windows.h>
#include <ShellScalingApi.h>
#pragma comment(lib, "shcore.lib")

#include "popup_window.h"
#include "keyboard_watcher.h"

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
  
  

  PopupWindow start_taskbar_popup([](HWND hwnd) {
    PAINTSTRUCT ps;
    HDC hdc = BeginPaint(hwnd, &ps);
    TextOut(hdc, 10, 5, "content here", 15);
    EndPaint(hwnd, &ps);
    return 0;
  });

  start_winkey_watcher(
    300,
    [&]{ start_taskbar_popup.show(100, 900, 200, 30); },
    [&]{ start_taskbar_popup.hide(); }
  );

  MSG msg;
  while (GetMessage(&msg, NULL, 0, 0)) {
    TranslateMessage(&msg);
    DispatchMessage(&msg);
  }
  return msg.wParam;
}