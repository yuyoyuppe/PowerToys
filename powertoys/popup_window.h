#pragma once
#include <functional>
#include <unordered_map>
#include <mutex>
#include <exception>
#include <windows.h>

typedef std::function<LRESULT(HWND hwnd)> PaintProc;

class PopupWindow {
public:
  PopupWindow(PaintProc paint_proc = PaintProc{}) {
    static const char* class_name = "PToyPopup";
    std::lock_guard<std::mutex> lock(static_mutex);
    if (!window_class_initialized) {
      WNDCLASSEX wnd_class;
      wnd_class.cbSize = sizeof(WNDCLASSEX);
      wnd_class.style = CS_SAVEBITS;
      wnd_class.lpfnWndProc = PopupWindow::window_proc;
      wnd_class.cbClsExtra = 0;
      wnd_class.cbWndExtra = 0;
      wnd_class.hInstance = GetModuleHandle(NULL);
      wnd_class.hIcon = LoadIcon(NULL, IDI_APPLICATION);
      wnd_class.hCursor = LoadCursor(NULL, IDC_ARROW);
      wnd_class.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
      wnd_class.lpszMenuName = NULL;
      wnd_class.lpszClassName = class_name;
      wnd_class.hIconSm = LoadIcon(NULL, IDI_APPLICATION);
      if (!RegisterClassEx(&wnd_class))
        throw std::runtime_error("Cannot register window class");
      window_class_initialized = true;
    }
    hwnd = CreateWindowEx(WS_EX_TOOLWINDOW | WS_EX_TOPMOST | WS_EX_LAYERED,
                          class_name, class_name,
                          WS_POPUP | WS_SYSMENU,
                          CW_USEDEFAULT,
                          CW_USEDEFAULT, 240, 120,
                          NULL, NULL,
                          GetModuleHandle(NULL),
                          NULL);
    if (hwnd == NULL)
      throw std::runtime_error("Cannot create window");
    SetLayeredWindowAttributes(hwnd, 0, (255 * 70) / 100, LWA_ALPHA);
    paint_procedures.emplace(hwnd, paint_proc);
  }
  
  void show(int x_pos, int y_pos, int x_size, int y_size, PaintProc paint_proc) {
    paint_procedures[hwnd] = paint_proc;
    show(x_pos, y_pos, x_size, y_size);
  }

  void show(int x_pos, int y_pos, int x_size, int y_size) {
    SetWindowPos(hwnd, HWND_TOPMOST, x_pos, y_pos, x_size, y_size, 0);
    HRGN elipse = CreateRoundRectRgn(0, 0, x_size, y_size, 15, 15);
    SetWindowRgn(hwnd, elipse, TRUE);
    show();
  }

  void show() {
    ShowWindow(hwnd, SW_SHOWNA);
    UpdateWindow(hwnd);
  }

  void hide() {
    ShowWindow(hwnd, SW_HIDE);
  }

  ~PopupWindow() {
    std::lock_guard<std::mutex> lock(static_mutex);
    paint_procedures.erase(hwnd);
  }

private:
  HWND hwnd;

  static LRESULT CALLBACK window_proc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_PAINT: {
      std::lock_guard<std::mutex> lock(static_mutex);
      std::unordered_map<HWND, PaintProc>::const_iterator iter = paint_procedures.find(hwnd);
      return iter != paint_procedures.end()
             ? iter->second(hwnd)
             : 0;
      }
    case WM_CLOSE:
      DestroyWindow(hwnd);
      break;
    case WM_DESTROY:
      PostQuitMessage(0);
      break;
    default:
      return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
  }
  static std::mutex static_mutex;
  static bool window_class_initialized;
  static std::unordered_map<HWND, PaintProc> paint_procedures;
};
