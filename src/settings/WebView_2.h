#pragma once
#include "StreamUriResolverFromFile.h"
#include <common/two_way_pipe_message_ipc.h>
#include <wrl.h>
#include <wil/com.h>
#include <WebView2.h>

#define SETTINGS_WEB_PAGE_LIGHT L"\\settings-html\\index.html"
#define SETTINGS_WEB_PAGE_DARK L"\\settings-html\\index-dark.html"

using namespace winrt::Windows::Web::UI::Interop;

typedef void (*PostMessageCallback)(const std::wstring&);

class WebView2Controller {
public:
  WebView2Controller(_In_ HWND hwnd, _In_ int nShowCmd, _In_ TwoWayPipeMessageIPC* messagePipe, _In_ bool darkMode, PostMessageCallback callback) {
    m_Hwnd = hwnd;
    m_ShowCmd = nShowCmd;
    m_MessagePipe = messagePipe;
    m_DarkMode = darkMode;
    m_PostMessageCallback = callback;
    InitializeWebView();
  }

  void PostData(const wchar_t* json_message);
  void ProcessExit();
  void Resize();

private:
  HWND m_Hwnd;
  int m_ShowCmd;
  wil::com_ptr<IWebView2WebView> m_WebView2;

  TwoWayPipeMessageIPC* m_MessagePipe;
  StreamUriResolverFromFile m_LocalUriResolver;
  bool m_DarkMode;
  PostMessageCallback m_PostMessageCallback;

  void InitializeWebView();
  void NavigateToSettingsWebPage();
};
