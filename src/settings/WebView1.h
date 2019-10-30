#pragma once

#include "StreamUriResolverFromFile.h"
#include <common/two_way_pipe_message_ipc.h>

#define URI_CONTENT_ID L"\\settings-html"
#define INDEX_LIGHT L"index.html"
#define INDEX_DARK L"index-dark.html"

using namespace winrt::Windows::Web::UI::Interop;

typedef void (*PostMessageCallback)(const std::wstring&);

class WebView1Controller {
public:
  WebView1Controller(_In_ HWND hwnd, _In_ int nShowCmd, _In_ TwoWayPipeMessageIPC* messagePipe, _In_ bool darkMode, PostMessageCallback callback) {
    m_Hwnd = hwnd;
    m_ShowCmd = nShowCmd;
    m_MessagePipe = messagePipe;
    m_DarkMode = darkMode;
    m_PostMessageCallback = callback;
    InitializeWebView();
  }

  void PostData(wchar_t* json_message);
  void ProcessExit();
  void Resize();

private:
  HWND m_Hwnd;
  int m_ShowCmd;
  WebViewControl m_WebView = nullptr;
  WebViewControlProcess m_WebViewProcess = nullptr;
  TwoWayPipeMessageIPC* m_MessagePipe;
  StreamUriResolverFromFile m_LocalUriResolver;
  bool m_DarkMode;
  PostMessageCallback m_PostMessageCallback;

  void InitializeWebView();
  void NavigateToLocalhostReactServer();
  void NavigateToUri(_In_ LPCWSTR uri_as_string);
};
