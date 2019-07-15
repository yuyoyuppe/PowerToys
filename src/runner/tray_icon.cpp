#include "pch.h"
#include "resource.h"
#include "settings_window.h"
#include "tray_icon.h"
#include <Windows.h>

extern "C" IMAGE_DOS_HEADER __ImageBase;

HWND tray_icon_hwnd = NULL;

// Message code that Windows will use for tray icon notifications.
UINT wm_icon_notify = 0;
UINT id_tray_icon = 0;

// Contains the Windows Message for taskbar creation.
UINT wm_taskbar_restart = 0;
UINT wm_run_on_main_ui_thread = 0;

// Struct to fill with callback and the data. The window_proc is responsible for cleaning it.
struct run_on_main_ui_thread_msg {
  main_loop_callback_function _callback;
  PVOID data;
};

bool dispatch_run_on_main_ui_thread(main_loop_callback_function _callback, PVOID data) {
  if (tray_icon_hwnd == NULL) {
    return false;
  }
  struct run_on_main_ui_thread_msg *wnd_msg = new struct run_on_main_ui_thread_msg();
  wnd_msg->_callback = _callback;
  wnd_msg->data = data;

  PostMessage(tray_icon_hwnd, wm_run_on_main_ui_thread, 0, (LPARAM)wnd_msg);

  return true;
}

NOTIFYICONDATA tray_icon_data;

LRESULT __stdcall tray_icon_window_proc(HWND window, UINT message, WPARAM wparam, LPARAM lparam) {
  switch (message) {
  case WM_CREATE:
    tray_icon_hwnd = window;
    wm_taskbar_restart = RegisterWindowMessage(TEXT("TaskbarCreated"));
    wm_run_on_main_ui_thread = RegisterWindowMessage(TEXT("RunOnMainThreadCallback"));
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
      case ID_SETTINGS_MENU_COMMAND:
        open_settings_window();
        break;
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
    } else if (message == wm_run_on_main_ui_thread) {
      if (lparam != NULL) {
        struct run_on_main_ui_thread_msg *msg = (struct run_on_main_ui_thread_msg *)lparam;
        msg->_callback(msg->data);
        delete msg;
        lparam = NULL;
      }
      break;
    }
  }
  return DefWindowProc(window, message, wparam, lparam);
}

void start_tray_icon() {
  id_tray_icon = wm_icon_notify = RegisterWindowMessage(TEXT("WM_PowerToysIconNotify"));

  auto h_instance = reinterpret_cast<HINSTANCE>(&__ImageBase);
  auto icon = LoadIcon(h_instance, MAKEINTRESOURCE(APPICON));

  static LPCTSTR class_name = TEXT("PToyTrayIconWindow");
  WNDCLASS wc = {};
  wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
  wc.hInstance = h_instance;
  wc.lpszClassName = class_name;
  wc.style = CS_HREDRAW | CS_VREDRAW;
  wc.lpfnWndProc = tray_icon_window_proc;
  wc.hIcon = icon;
  RegisterClass(&wc);
  auto hwnd = CreateWindow(wc.lpszClassName,
                           TEXT("PToyTrayIconWindow"),
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
  wcscpy_s(tray_icon_data.szTip, sizeof(tray_icon_data.szTip), TEXT("PowerToys"));
  tray_icon_data.uFlags = NIF_ICON | NIF_TIP | NIF_MESSAGE;

  Shell_NotifyIcon(NIM_ADD, &tray_icon_data);
}
