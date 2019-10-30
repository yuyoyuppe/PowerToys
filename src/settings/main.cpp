#include "pch.h"
#include <Commdlg.h>
#include "StreamUriResolverFromFile.h"
#include <Shellapi.h>
#include <common/two_way_pipe_message_ipc.h>
#include <ShellScalingApi.h>
#include "resource.h"
#include <common/dpi_aware.h>
#include <common/common.h>
#include "WebView1.h"
#include "WindowUserMessages.h"

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "shcore.lib")
#pragma comment(lib, "windowsapp")

#ifdef _DEBUG
#define _DEBUG_WITH_LOCALHOST 0
// Define as 1 For debug purposes, to access localhost servers.
// webview_process_options.PrivateNetworkClientServerCapability(winrt::Windows::Web::UI::Interop::WebViewControlProcessCapabilityState::Enabled);
// To access localhost:8080 for development, you'll also need to disable loopback restrictions for the webview:
// > checknetisolation LoopbackExempt -a -n=Microsoft.Win32WebViewHost_cw5n1h2txyewy
// To remove the exception after development:
// > checknetisolation LoopbackExempt -d -n=Microsoft.Win32WebViewHost_cw5n1h2txyewy
// Source: https://github.com/windows-toolkit/WindowsCommunityToolkit/issues/2226#issuecomment-396360314
#endif

HINSTANCE g_hinst = nullptr;
HWND g_hwnd = nullptr;
WebView1Controller* g_webview_controller = nullptr;

// Message pipe to send/receive messages to/from the Powertoys runner.
TwoWayPipeMessageIPC* g_message_pipe = nullptr;

// Set to true if waiting for webview confirmation before closing the Window.
bool g_waiting_for_exit_confirmation = false;

// Is the setting window to be started in dark mode
bool g_start_in_dark_mode = false;

#define SEND_TO_WEBVIEW_MSG 1

// message_pipe_callback reveives the messages sent by the PowerToys runner and
// dispatches them to the WebView control.
void message_pipe_callback(const std::wstring& msg) {
  if (g_hwnd != nullptr) {
    // Allocate the COPYDATASTRUCT and message to pass to the Webview.
    // This is needed in order to use PostMessage, since COM calls to
    // g_webview.InvokeScriptAsync can't be made from other threads.

    PCOPYDATASTRUCT message = new COPYDATASTRUCT();
    DWORD buff_size = (DWORD)(msg.length() + 1);

    // 'wnd_static_proc()' will free the buffer allocated here.
    wchar_t* buffer = new wchar_t[buff_size];

    wcscpy_s(buffer, buff_size, msg.c_str());
    message->dwData = SEND_TO_WEBVIEW_MSG;
    message->cbData = buff_size * sizeof(wchar_t);
    message->lpData = (PVOID)buffer;
    WINRT_VERIFY(PostMessage(g_hwnd, WM_USER_POST_TO_WEBVIEW, (WPARAM)g_hwnd, (LPARAM)message));
  }
}

void send_message_to_powertoys_runner(_In_ const std::wstring& msg) {
  if (g_message_pipe != nullptr) {
    g_message_pipe->send(msg);
  } else {
#ifdef _DEBUG
#include "DebugData.h"
    // If PowerToysSettings was not started by the PowerToys runner, simulate
    // the data as if it was sent by the runner.
    MessageBox(nullptr, msg.c_str(), L"From Webview", MB_OK);
    message_pipe_callback(debug_settings_info);
#endif
  }
}

// Called by the WebView control.
// The message can be:
// 1 - json data to be forward to the runner
// 2 - exit confirmation by the user
// 3 - exit canceled by the user
void webview_message_callback(const std::wstring& msg) {
  if (msg[0] == '{') {
    // It's a JSON string, send the message to the PowerToys runner.
    std::thread(send_message_to_powertoys_runner, msg).detach();
  } else {
    // It's not a JSON string, check for expected control messages.
    if (msg == L"exit") {
      // WebView confirms the settings application can exit.
      WINRT_VERIFY(PostMessage(g_hwnd, WM_USER_EXIT, 0, 0));
    } else if (msg == L"cancel-exit") {
      // WebView canceled the exit request.
      g_waiting_for_exit_confirmation = false;
    }
  }
}

LRESULT CALLBACK wnd_proc_static(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
  switch (message) {
  case WM_CLOSE:
    if (g_waiting_for_exit_confirmation) {
      // If another WM_CLOSE is received while waiting for webview confirmation,
      // allow DefWindowProc to be called and destroy the window.
      break;
    } else {
      // Allow user to confirm exit in the WebView in case there's possible data loss.
      g_waiting_for_exit_confirmation = true;
      if (g_webview_controller != nullptr) {
        g_webview_controller->ProcessExit();
        return 0;
      }
    }
    break;
  case WM_DESTROY:
    PostQuitMessage(0);
    break;
  case WM_SIZE:
    if (g_webview_controller != nullptr) {
      g_webview_controller->Resize();
    }
    break;
  case WM_CREATE:
    break;
  case WM_DPICHANGED:
  {
    // Resize the window using the suggested rect
    RECT* const prcNewWindow = (RECT*)lParam;
    SetWindowPos(hWnd,
      nullptr,
      prcNewWindow->left,
      prcNewWindow->top,
      prcNewWindow->right - prcNewWindow->left,
      prcNewWindow->bottom - prcNewWindow->top,
      SWP_NOZORDER | SWP_NOACTIVATE);
  }
    break;
  case WM_NCCREATE:
  {
    // Enable auto-resizing the title bar
    EnableNonClientDpiScaling(hWnd);
  }
    break;
  case WM_USER_POST_TO_WEBVIEW:
  {
    PCOPYDATASTRUCT msg = (PCOPYDATASTRUCT)lParam;
    if (msg->dwData == SEND_TO_WEBVIEW_MSG) {
      wchar_t* json_message = (wchar_t*)(msg->lpData);
      if (g_webview_controller != nullptr) {
        g_webview_controller->PostData(json_message);
      }
      delete[] json_message;
    }
    // wnd_proc_static is responsible for freeing memory.
    delete msg;
  }
    break;
  case WM_USER_EXIT:
    DestroyWindow(hWnd);
    break;
  default:
    break;
  }

  return DefWindowProc(hWnd, message, wParam, lParam);;
}

void register_classes(HINSTANCE hInstance) {
  WNDCLASSEXW wcex;
  wcex.cbSize = sizeof(WNDCLASSEX);

  wcex.style = 0;
  wcex.lpfnWndProc = wnd_proc_static;
  wcex.cbClsExtra = 0;
  wcex.cbWndExtra = 0;
  wcex.hInstance = hInstance;
  wcex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(APPICON));
  wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
  wcex.hbrBackground = g_start_in_dark_mode ? CreateSolidBrush(0) : (HBRUSH)(COLOR_WINDOW + 1);
  wcex.lpszMenuName = nullptr;
  wcex.lpszClassName = L"PTSettingsClass";
  wcex.hIconSm = nullptr;

  WINRT_VERIFY(RegisterClassExW(&wcex));
}

HWND create_main_window(HINSTANCE hInstance) {
  RECT desktopRect;
  const HWND hDesktop = GetDesktopWindow();
  WINRT_VERIFY(hDesktop != nullptr);
  WINRT_VERIFY(GetWindowRect(hDesktop, &desktopRect));

  int wind_width = 1024;
  int wind_height = 700;
  DPIAware::Convert(nullptr, wind_width, wind_height);

  return CreateWindowW(
    L"PTSettingsClass",
    L"PowerToys Settings",
    WS_OVERLAPPEDWINDOW,
    (desktopRect.right - wind_width)/2,
    (desktopRect.bottom - wind_height)/2,
    wind_width,
    wind_height,
    nullptr,
    nullptr,
    hInstance,
    nullptr);
}

void wait_on_parent_process(DWORD pid) {
  HANDLE process = OpenProcess(SYNCHRONIZE, FALSE, pid);
  if (process != nullptr) {
    if (WaitForSingleObject(process, INFINITE) == WAIT_OBJECT_0) {
      // If it's possible to detect when the PowerToys process terminates, message the main window.
      CloseHandle(process);
      if (g_hwnd) {
        WINRT_VERIFY(PostMessage(g_hwnd, WM_USER_EXIT, 0, 0));
      }
    } else {
      CloseHandle(process);
    }
  }
}

// Parse arguments: initialize two-way IPC message pipe and if settings window is to be started
// in dark mode.
void parse_args() {
  // Expected calling arguments:
  // [0] - This executable's path.
  // [1] - PowerToys pipe server.
  // [2] - Settings pipe server.
  // [3] - PowerToys process pid.
  // [4] - optional "dark" parameter if the settings window is to be started in dark mode
  LPWSTR *argument_list;
  int n_args;

  argument_list = CommandLineToArgvW(GetCommandLineW(), &n_args);
  if (n_args > 3) {
    g_message_pipe = new TwoWayPipeMessageIPC(std::wstring(argument_list[2]), std::wstring(argument_list[1]), message_pipe_callback);
    g_message_pipe->start(nullptr);

    DWORD parent_pid = std::stol(argument_list[3]);
    std::thread(wait_on_parent_process, parent_pid).detach();
  } else {
#ifndef _DEBUG
    MessageBox(nullptr, L"This executable isn't supposed to be called as a stand-alone process", L"Error running settings", MB_OK);
    exit(1);
#endif
  }

  if (n_args > 4) {
    g_start_in_dark_mode = wcscmp(argument_list[4], L"dark") == 0;
  }

  LocalFree(argument_list);
}

int WINAPI WinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPSTR lpCmdLine, _In_ int nShowCmd) {
  CoInitialize(nullptr);

  if (is_process_elevated()) {
    if (!drop_elevated_privileges()) {
      MessageBox(NULL, L"Failed to drop admin privileges.\nPlease report the bug to https://github.com/microsoft/PowerToys/issues", L"PowerToys Settings Error", MB_OK);
      exit(1);
    }
  }

  g_hinst = hInstance;
  parse_args();
  register_classes(hInstance);
  g_hwnd = create_main_window(hInstance);
  if (g_hwnd == nullptr) {
    MessageBox(NULL, L"Failed to create main window.\nPlease report the bug to https://github.com/microsoft/PowerToys/issues", L"PowerToys Settings Error", MB_OK);
    exit(1);
  }
  g_webview_controller = new WebView1Controller(g_hwnd, nShowCmd, g_message_pipe, g_start_in_dark_mode, webview_message_callback);

  // Main message loop.
  MSG msg;
  while (GetMessage(&msg, nullptr, 0, 0) !=0 ) {
    TranslateMessage(&msg);
    DispatchMessage(&msg);
  }

  return (int)msg.wParam;
}
