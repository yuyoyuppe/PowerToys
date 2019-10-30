#include "pch.h"
#include "WebView_2.h"
#include <Commdlg.h>
#include "StreamUriResolverFromFile.h"
#include <Shellapi.h>
#include <ShellScalingApi.h>
#include "resource.h"
#include <common/dpi_aware.h>
#include <common/common.h>
#include <wrl.h>
#include <wil/com.h>

using namespace winrt;
using namespace winrt::Windows::Foundation;
using namespace winrt::Windows::Web::UI;
using namespace winrt::Windows::Web::UI::Interop;
using namespace Microsoft::WRL;

void WebView2Controller::PostData(const wchar_t* json_message) {
  if (m_WebView2 != nullptr) {
    m_WebView2->PostWebMessageAsJson(json_message);
  }
}

void WebView2Controller::ProcessExit() {
  // TODO
}

void WebView2Controller::Resize() {
  if (m_WebView2 != nullptr) {
    RECT bounds;
    GetClientRect(m_Hwnd, &bounds);
    m_WebView2->put_Bounds(bounds);
  }
}

void WebView2Controller::NavigateToSettingsWebPage() {
  WCHAR path[MAX_PATH];
  WINRT_VERIFY(GetModuleFileName(NULL, path, MAX_PATH));
  WINRT_VERIFY(PathRemoveFileSpec(path));
  if (m_DarkMode) {
    wcscat_s(path, SETTINGS_WEB_PAGE_DARK);
  } else {
    wcscat_s(path, SETTINGS_WEB_PAGE_LIGHT);
  }
  HRESULT hres = m_WebView2->Navigate(path);
}

void WebView2Controller::InitializeWebView() {
  CreateWebView2Environment(Callback<IWebView2CreateWebView2EnvironmentCompletedHandler>(
    [this](HRESULT result, IWebView2Environment* env) -> HRESULT {
      // Create a WebView, whose parent is the main window hwnd
      env->CreateWebView(m_Hwnd, Callback<IWebView2CreateWebViewCompletedHandler>(
        [this](HRESULT result, IWebView2WebView* webview) -> HRESULT {
          if (webview != nullptr) {
            m_WebView2 = webview;
          }

          // Default settings are:
          // Settings->put_IsScriptEnabled(TRUE);
          // Settings->put_AreDefaultScriptDialogsEnabled(TRUE);
          // Settings->put_IsWebMessageEnabled(TRUE);

          IWebView2Settings* Settings;
          m_WebView2->get_Settings(&Settings);

          // Resize WebView to fit the bounds of the parent window
          Resize();

          // Schedule an async task to navigate to the settings general page
          NavigateToSettingsWebPage();

          // Step 4 - Navigation events

          // Step 5 - Scripting

          // Step 6 - Communication between host and web content
          EventRegistrationToken token;
          m_WebView2->add_WebMessageReceived(Callback<IWebView2WebMessageReceivedEventHandler>(
            [this](IWebView2WebView* webview, IWebView2WebMessageReceivedEventArgs* args) -> HRESULT {
              PWSTR message;
              args->get_WebMessageAsString(&message);
              m_PostMessageCallback(message);
              CoTaskMemFree(message);
              return S_OK;
            }).Get(), &token);

          return S_OK;
        }).Get());
      return S_OK;
    }).Get());
}
