#include "start_visible.h"
#include "com.h"
#include <Windows.h>
#include <objbase.h>
#include <shobjidl.h>
#include <wrl.h>
#include <stdexcept>

bool is_start_visible_impl() {
  static Microsoft::WRL::ComPtr<IAppVisibility> app_visibility;
  if (app_visibility.Get() == nullptr) {
    auto hr = CoCreateInstance(CLSID_AppVisibility,
                               nullptr,
                               CLSCTX_INPROC_SERVER,
                               IID_PPV_ARGS(&app_visibility));
    if (!SUCCEEDED(hr)) {
      throw std::runtime_error("Cannot create IAppVisibility");
    }
  }
  BOOL visible;
  auto result = app_visibility->IsLauncherVisible(&visible);
  return SUCCEEDED(result) && visible;
}

bool is_start_visible() {
  // make sure CoUninitilize will be called after
  // app_visibility is destroyed
  init_com();
  return is_start_visible_impl();
}

void init_start_visible() {
  // is_start_visible lazy loads Start menu watcher,
  // let us call it to initialize it
  is_start_visible();
}