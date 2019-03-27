#include <sstream>
#include <windows.h>
#include <ShellScalingApi.h>
#pragma comment(lib, "shcore.lib")

#include "popup_window.h"
#include "keyboard_watcher.h"
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
  
  

  PopupWindow start_taskbar_popup([](HWND hwnd) {
    PAINTSTRUCT ps;
    HDC hdc = BeginPaint(hwnd, &ps);
    TextOut(hdc, 10, 5, "content here", 15);
    EndPaint(hwnd, &ps);
    return 0;
  });

  start_winkey_watcher(
    300,
    [&]{
      HWND active_window = GetForegroundWindow();
      if (active_window == NULL) {
        return;
      }
      auto dpi = GetDpiForWindow(active_window);
      auto max_button = getWindowMaximizeButton(active_window);
      RECT rect;
      if (GetWindowRect(active_window, &rect) == 0) {
        return;
      }
      start_taskbar_popup.show(
        rect.left + (rect.right - rect.left) / 2,
        rect.top + 150,
        600, 30,
        [=](HWND hwnd) {
          PAINTSTRUCT ps;
          HDC hdc = BeginPaint(hwnd, &ps);
          std::stringstream stream;
          stream << "HWND: " << active_window
                 << " DPI: " << dpi;
          if (max_button) {
            stream << " MaxButton: (" << max_button->top << ", " << max_button->left << ")x("
                                      << max_button->bottom << ", " << max_button->right << ")";
          } else {
            stream << " MaxButton: Unknow";
          }
          auto str = stream.str();
          TextOut(hdc, 10, 5, str.c_str(), str.length());
          EndPaint(hwnd, &ps);

          if (max_button) {
            HDC hDC_Desktop = GetDC(0);
            HBRUSH blueBrush = CreateSolidBrush(RGB(0, 0, 255));
            FillRect(hDC_Desktop, &max_button.value(), blueBrush);
            ReleaseDC(0, hDC_Desktop);
          }
          return 0;
        }
      );
      SetForegroundWindow(active_window);
    },
    [&]{ start_taskbar_popup.hide(); }
  );

  MSG msg;
  while (GetMessage(&msg, NULL, 0, 0)) {
    TranslateMessage(&msg);
    DispatchMessage(&msg);
  }
  return msg.wParam;
}