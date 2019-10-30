#include "pch.h"
#include "WebView1.h"
#include <Commdlg.h>
#include "StreamUriResolverFromFile.h"
#include <Shellapi.h>
#include <ShellScalingApi.h>
#include "resource.h"
#include <common/dpi_aware.h>
#include <common/common.h>

using namespace winrt;
using namespace winrt::Windows::Foundation;
using namespace winrt::Windows::Web::UI;
using namespace winrt::Windows::Web::UI::Interop;

#ifdef _DEBUG
void WebView1Controller::NavigateToLocalhostReactServer() {
  // Useful for connecting to instance running in react development server.
  m_WebView.Navigate(Uri(hstring(L"http://localhost:8080")));
}
#endif

void WebView1Controller::NavigateToUri(_In_ LPCWSTR uri_as_string) {
  // initialize the base_path for the html content relative to the executable.
  WINRT_VERIFY(GetModuleFileName(nullptr, m_LocalUriResolver.base_path, MAX_PATH));
  WINRT_VERIFY(PathRemoveFileSpec(m_LocalUriResolver.base_path));
  wcscat_s(m_LocalUriResolver.base_path, URI_CONTENT_ID);
  Uri url = m_WebView.BuildLocalStreamUri(hstring(URI_CONTENT_ID), hstring(uri_as_string));
  m_WebView.NavigateToLocalStreamUri(url, m_LocalUriResolver);
}

Rect client_rect_to_bounds_rect(_In_ HWND hwnd) {
  RECT client_rect = { 0 };
  WINRT_VERIFY(GetClientRect(hwnd, &client_rect));

  Rect bounds =
  {
    0,
    0,
    static_cast<float>(client_rect.right - client_rect.left),
    static_cast<float>(client_rect.bottom - client_rect.top)
  };

  return bounds;
}

void WebView1Controller::Resize() {
  Rect bounds = client_rect_to_bounds_rect(m_Hwnd);
  IWebViewControlSite webViewControlSite = (IWebViewControlSite)m_WebView;
  webViewControlSite.Bounds(bounds);
}

void WebView1Controller::PostData(wchar_t* json_message) {
  const auto _ = m_WebView.InvokeScriptAsync(hstring(L"receive_from_settings_app"), { hstring(json_message) });
}

void WebView1Controller::ProcessExit() {
  const auto _ = m_WebView.InvokeScriptAsync(hstring(L"exit_settings_app"), {});
}

extern void process_message_from_webview(const std::wstring& msg);

void WebView1Controller::InitializeWebView() {
  try {
    m_WebViewProcess = WebViewControlProcess();
    auto asyncwebview = m_WebViewProcess.CreateWebViewControlAsync((int64_t)m_Hwnd, client_rect_to_bounds_rect(m_Hwnd));
    asyncwebview.Completed([=](IAsyncOperation<WebViewControl> const& sender, AsyncStatus status) {
      if (status == AsyncStatus::Completed) {
        WINRT_VERIFY(sender != nullptr);
        m_WebView = sender.GetResults();
        WINRT_VERIFY(m_WebView != nullptr);

        // In order to receive window.external.notify() calls in ScriptNotify
        m_WebView.Settings().IsScriptNotifyAllowed(true);

        m_WebView.Settings().IsJavaScriptEnabled(true);

        m_WebView.NewWindowRequested([=](IWebViewControl sender_requester, WebViewControlNewWindowRequestedEventArgs args) {
          // Open the requested link in the default browser registered in the Shell
          int res = static_cast<int>(reinterpret_cast<uintptr_t>(ShellExecute(nullptr, L"open", args.Uri().AbsoluteUri().c_str(), nullptr, nullptr, SW_SHOWNORMAL)));
          WINRT_VERIFY(res > 32);
          });

        m_WebView.DOMContentLoaded([=](IWebViewControl sender_loaded, WebViewControlDOMContentLoadedEventArgs const& args_loaded) {
          ShowWindow(m_Hwnd, m_ShowCmd);
          });

        m_WebView.ScriptNotify([=](IWebViewControl sender_script_notify, WebViewControlScriptNotifyEventArgs const& args_script_notify) {
          // WebView called window.external.notify()
          std::wstring message = args_script_notify.Value().c_str();

          m_PostMessageCallback(message);
          
          });

        m_WebView.AcceleratorKeyPressed([&](IWebViewControl sender, WebViewControlAcceleratorKeyPressedEventArgs const& args) {
          if (args.VirtualKey() == winrt::Windows::System::VirtualKey::F4) {
            // WebView swallows key-events. Detect Alt-F4 one and close the window manually.
            ProcessExit();
          }
          });

        Resize();

#if defined(_DEBUG) && _DEBUG_WITH_LOCALHOST
        // Navigates to localhost:8080
        NavigateToLocalhostReactServer();
#else
        // Navigates to settings-html/index.html or index-dark.html
        NavigateToUri(m_DarkMode ? INDEX_DARK : INDEX_LIGHT);
#endif
      } else if (status == AsyncStatus::Error) {
        MessageBox(NULL, L"Failed to create the WebView control.\nPlease report the bug to https://github.com/microsoft/PowerToys/issues", L"PowerToys Settings Error", MB_OK);
        exit(1);
      } else if (status == AsyncStatus::Started) {
        // Ignore.
      } else if (status == AsyncStatus::Canceled) {
        // Ignore.
      }
      });
  } catch (hresult_error const& e) {
    WCHAR message[1024] = L"";
    StringCchPrintf(message, ARRAYSIZE(message), L"failed: %ls", e.message().c_str());
    MessageBox(m_Hwnd, message, L"Error", MB_OK);
  }
}
