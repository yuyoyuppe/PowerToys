#include "pch.h"
#include <Commdlg.h>
#include "StreamUriResolverFromFile.h"
#include <Shellapi.h>

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "shcore.lib")
#pragma comment(lib, "windowsapp")
#pragma comment(lib, "dxgi")
#pragma comment(lib, "d3d11")
#pragma comment(lib, "d2d1")
#pragma comment(lib, "dcomp")
#pragma comment(lib, "dwmapi")

HINSTANCE m_hInst;
HWND main_window_handler = nullptr;
using namespace winrt;
using namespace winrt::Windows::Foundation;
using namespace winrt::Windows::Storage::Streams;
using namespace winrt::Windows::Web::Http;
using namespace winrt::Windows::Web::Http::Headers;
using namespace winrt::Windows::Web::UI;
using namespace winrt::Windows::Web::UI::Interop;
using namespace winrt::Windows::System;

winrt::Windows::Web::UI::Interop::WebViewControl webview_control = nullptr;
winrt::Windows::Web::UI::Interop::WebViewControlProcess webview_process = nullptr;
winrt::Windows::Web::UI::Interop::WebViewControlProcessOptions webview_process_options = nullptr;
StreamUriResolverFromFile local_uri_resolver;
//Microsoft::WRL::ComPtr<ABI::Windows::Web::IUriToStreamResolver> coiso;



void NavigateToUri(_In_ LPCWSTR uriAsString) {
  Uri url = webview_control.BuildLocalStreamUri(hstring(L"settings-html"), hstring(uriAsString));
  //m_webViewControl.Navigate(Uri(hstring(uriAsString)));
  webview_control.NavigateToLocalStreamUri(url, local_uri_resolver);

}

Rect hwnd_client_rect_to_bounds_rect(_In_ HWND hwnd) {
  RECT client_rect = { 0 };
  GetClientRect(hwnd, &client_rect);

  Rect bounds =
  {
    0,
    0,
    static_cast<float>(client_rect.right - client_rect.left),
    static_cast<float>(client_rect.bottom - client_rect.top)
  };

  return bounds;
}

void resize_web_view() {
  Rect bounds = hwnd_client_rect_to_bounds_rect(main_window_handler);
  winrt::Windows::Web::UI::Interop::IWebViewControlSite webViewControlSite = (winrt::Windows::Web::UI::Interop::IWebViewControlSite) webview_control;
  webViewControlSite.Bounds(bounds);

}

void initialize_win32_webview() {
  
  // initialize the base_path for the html content relative to the executable.
  TCHAR executable_path[MAX_PATH];
  GetModuleFileName(NULL, executable_path, MAX_PATH);
  PathRemoveFileSpec(executable_path);
  wcscat_s(executable_path, L"\\settings-html");
  wcscpy_s(local_uri_resolver.base_path, executable_path);
  
  try {
    if (!webview_process_options) {
      webview_process_options = winrt::Windows::Web::UI::Interop::WebViewControlProcessOptions();
    }

    if (!webview_process) {
      webview_process = winrt::Windows::Web::UI::Interop::WebViewControlProcess();
    }

    auto asyncwebview = webview_process.CreateWebViewControlAsync((int64_t)main_window_handler, hwnd_client_rect_to_bounds_rect(main_window_handler));
    asyncwebview.Completed([=](IAsyncOperation<WebViewControl> const& sender, AsyncStatus args) {
      webview_control = sender.GetResults();
      // In order to receive window.external.notify() calls in ScriptNotify
      webview_control.Settings().IsScriptNotifyAllowed(true);
      webview_control.Settings().IsJavaScriptEnabled(true);
      webview_control.DOMContentLoaded([=](IWebViewControl sender_loaded, WebViewControlDOMContentLoadedEventArgs const& args_loaded) {
        //auto scriptargs = { hstring(L"window.external.notify('test');") };
        //m_webViewControl.InvokeScriptAsync(hstring(L"eval"), scriptargs);
      });
      webview_control.ScriptNotify([=](IWebViewControl sender_script_notify, WebViewControlScriptNotifyEventArgs const& args_script_notify) {
        std::wstring message_sent = args_script_notify.Value().c_str();
        MessageBox(main_window_handler, message_sent.c_str(), L"Message from WebView", MB_OK);
      });
      resize_web_view();
      NavigateToUri(L"index.html");
    });
  }
  catch (hresult_error const& e) {
    TCHAR message[1024] = L"";
    StringCchPrintf(message, ARRAYSIZE(message), L"failed: %ls", e.message().c_str());
    MessageBox(main_window_handler, message, L"Error", MB_OK);
  }
}



LRESULT CALLBACK wnd_proc_static(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
  switch (message) {
  case WM_DESTROY:
    PostQuitMessage(0);
    break;
  case WM_SIZE:
    if (webview_control != nullptr) {
      resize_web_view();
    }
    break;
  }
  return DefWindowProc(hWnd, message, wParam, lParam);;
}

void register_classes(HINSTANCE hInstance) {
  WNDCLASSEXW wcex;
  wcex.cbSize = sizeof(WNDCLASSEX);

  wcex.style = CS_HREDRAW | CS_VREDRAW;
  wcex.lpfnWndProc = wnd_proc_static;
  wcex.cbClsExtra = 0;
  wcex.cbWndExtra = 0;
  wcex.hInstance = hInstance;
  wcex.hIcon = nullptr;
  wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
  wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
  wcex.lpszMenuName = nullptr;
  wcex.lpszClassName = L"PTSettingsClass";
  wcex.hIconSm = nullptr;

  RegisterClassExW(&wcex);
}

int init_instance(HINSTANCE hInstance, int nCmdShow) {
  m_hInst = hInstance;
  main_window_handler = CreateWindow(TEXT("PTSettingsClass"), TEXT("PowerToys Settings"), WS_OVERLAPPEDWINDOW,
    CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

  initialize_win32_webview();
  ShowWindow(main_window_handler, nCmdShow);
  UpdateWindow(main_window_handler);

  return TRUE;
}

int start_webview_window(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
  register_classes(hInstance);
  init_instance(hInstance, nCmdShow);
  MSG msg;
  // Main message loop:
  while (GetMessage(&msg, nullptr, 0, 0)) {
    TranslateMessage(&msg);
    DispatchMessage(&msg);
  }

  return (int)msg.wParam;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
  HRESULT hrInit = CoInitialize(nullptr);
  return start_webview_window(hInstance, hPrevInstance, lpCmdLine, nCmdShow);
}

