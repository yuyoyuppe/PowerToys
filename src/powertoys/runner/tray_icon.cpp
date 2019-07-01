#include "pch.h"
#include "functionalities.h"
#include "resource.h"
#include <Windows.h>

extern "C" IMAGE_DOS_HEADER __ImageBase;

// Message code that Windows will use for tray icon notifications.
UINT wm_icon_notify = 0;
UINT id_tray_icon = 0;

// Contais the Windows Message for taskbar creation.
UINT wm_taskbar_restart = 0;

NOTIFYICONDATA tray_icon_data;

LRESULT __stdcall tray_icon_window_proc(HWND window, UINT message, WPARAM wparam, LPARAM lparam) {
  switch (message) {
  case WM_CREATE:
    wm_taskbar_restart = RegisterWindowMessage(TEXT("TaskbarCreated"));
    break;
  case WM_DESTROY:
    Shell_NotifyIcon(NIM_DELETE, &tray_icon_data);
    PostQuitMessage(0);
    break;
  case WM_CLOSE:
    DestroyWindow(window);
    break;
  case WM_COMMAND:
    switch(wparam) {
      case ID_EXIT_MENU_COMMAND:
        DestroyWindow(window);
        break;
    }
    break;
  default:
    if (message == wm_icon_notify) {
      switch(lparam) {
        case WM_RBUTTONUP:
        case WM_CONTEXTMENU:
        {
          auto h_menu = LoadMenu(reinterpret_cast<HINSTANCE>(&__ImageBase), MAKEINTRESOURCE(ID_TRAY_MENU));
          auto h_sub_menu = GetSubMenu(h_menu, 0);
          POINT mouse_pointer;
          GetCursorPos(&mouse_pointer);
          SetForegroundWindow(window); // Needed for the context menu to disappear.
          TrackPopupMenu(h_sub_menu, TPM_CENTERALIGN|TPM_BOTTOMALIGN, mouse_pointer.x, mouse_pointer.y, 0, window, nullptr);
          DestroyMenu(h_menu);
        }
        break;
      }
    } else if(message == wm_taskbar_restart) {
      // To show tray icon when the taskbar is created/restarted.
      Shell_NotifyIcon(NIM_ADD, &tray_icon_data);
    }
  }
  return DefWindowProc(window, message, wparam, lparam);
}

void start_tray_icon() {
  id_tray_icon = wm_icon_notify = RegisterWindowMessage("WM_PowerToysIconNotify");

  auto h_instance = reinterpret_cast<HINSTANCE>(&__ImageBase);
  auto icon = LoadIcon(h_instance, MAKEINTRESOURCE(APPICON));

  static const char* class_name = "PToyTrayIconWindow";
  WNDCLASS wc = {};
  wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
  wc.hInstance = h_instance;
  wc.lpszClassName = class_name;
  wc.style = CS_HREDRAW | CS_VREDRAW;
  wc.lpfnWndProc = tray_icon_window_proc;
  wc.hIcon = icon;
  RegisterClass(&wc);
  auto hwnd = CreateWindow(wc.lpszClassName,
                           "PToyTrayIconWindow",
                           WS_OVERLAPPEDWINDOW | WS_POPUP,
                           CW_USEDEFAULT,
                           CW_USEDEFAULT,
                           CW_USEDEFAULT,
                           CW_USEDEFAULT,
                           nullptr,
                           nullptr,
                           wc.hInstance,
                           nullptr);
  WINRT_VERIFY(hwnd);

  memset(&tray_icon_data, 0, sizeof(tray_icon_data));
  tray_icon_data.cbSize = sizeof(tray_icon_data);
  tray_icon_data.hIcon = icon;
  tray_icon_data.hWnd = hwnd;
  tray_icon_data.uID = id_tray_icon;
  tray_icon_data.uCallbackMessage = wm_icon_notify;
  strcpy_s(tray_icon_data.szTip, sizeof(tray_icon_data.szTip), "PowerToys");
  tray_icon_data.uFlags = NIF_ICON | NIF_TIP | NIF_MESSAGE;

  Shell_NotifyIcon(NIM_ADD, &tray_icon_data);
}
