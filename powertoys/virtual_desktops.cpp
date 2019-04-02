#include "virtual_desktops.h"
#include <objbase.h>
#include <ObjectArray.h>
#include <ShObjIdl_core.h>
#include <wrl.h>
#include <stdexcept>
#include "com.h"

const CLSID CLSID_ImmersiveShell = { 0xC2F03A33, 0x21F5, 0x47FA, 0xB4, 0xBB, 0x15, 0x63, 0x62, 0xA2, 0xF2, 0x39 };
const CLSID CLSID_VirtualDesktopAPI_Unknown = { 0xC5E0CDCA, 0x7B6E, 0x41B2, 0x9F, 0xC4, 0xD9, 0x39, 0x75, 0xCC, 0x46, 0x7B };
const IID IID_IVirtualDesktopManagerInternal = { 0xF31574D6, 0xB682, 0x4CDC, 0xBD, 0x56, 0x18, 0x27, 0x86, 0x0A, 0xBE, 0xC6 };
const CLSID CLSID_IVirtualNotificationService = { 0xA501FDEC, 0x4A09, 0x464C, 0xAE, 0x4E, 0x1B, 0x9C, 0x21, 0xB8, 0x49, 0x18 };

struct IApplicationView : public IUnknown { };

EXTERN_C const IID IID_IVirtualDesktop;
MIDL_INTERFACE("FF72FFDD-BE7E-43FC-9C03-AD81681E88E4") IVirtualDesktop : public IUnknown{
public:
  virtual HRESULT STDMETHODCALLTYPE IsViewVisible(IApplicationView *pView, int *pfVisible) = 0;
  virtual HRESULT STDMETHODCALLTYPE GetID(GUID *pGuid) = 0;
};

enum AdjacentDesktop
{
  LeftDirection = 3,
  RightDirection = 4
};

EXTERN_C const IID IID_IVirtualDesktopManagerInternal;
MIDL_INTERFACE("F31574D6-B682-4CDC-BD56-1827860ABEC6") IVirtualDesktopManagerInternal : public IUnknown{
public:
  virtual HRESULT STDMETHODCALLTYPE GetCount(UINT *pCount) = 0;
  virtual HRESULT STDMETHODCALLTYPE MoveViewToDesktop(IApplicationView *pView, IVirtualDesktop *pDesktop) = 0;
  virtual HRESULT STDMETHODCALLTYPE CanViewMoveDesktops(IApplicationView *pView, int *pfCanViewMoveDesktops) = 0;
  virtual HRESULT STDMETHODCALLTYPE GetCurrentDesktop(IVirtualDesktop** desktop) = 0;
  virtual HRESULT STDMETHODCALLTYPE GetDesktops(IObjectArray **ppDesktops) = 0;
  virtual HRESULT STDMETHODCALLTYPE GetAdjacentDesktop(IVirtualDesktop *pDesktopReference, AdjacentDesktop uDirection, IVirtualDesktop **ppAdjacentDesktop) = 0;
  virtual HRESULT STDMETHODCALLTYPE SwitchDesktop(IVirtualDesktop *pDesktop) = 0;
  virtual HRESULT STDMETHODCALLTYPE CreateDesktopW(IVirtualDesktop **ppNewDesktop) = 0;
  virtual HRESULT STDMETHODCALLTYPE RemoveDesktop(IVirtualDesktop *pRemove, IVirtualDesktop *pFallbackDesktop) = 0;
  virtual HRESULT STDMETHODCALLTYPE FindDesktop(GUID *desktopId, IVirtualDesktop **ppDesktop) = 0;
};

/*
EXTERN_C const IID IID_IVirtualDesktopManager;
MIDL_INTERFACE("a5cd92ff-29be-454c-8d04-d82879fb3f1b") IVirtualDesktopManager : public IUnknown{
public:
  virtual HRESULT STDMETHODCALLTYPE IsWindowOnCurrentVirtualDesktop(__RPC__in HWND topLevelWindow, __RPC__out BOOL *onCurrentDesktop) = 0;
  virtual HRESULT STDMETHODCALLTYPE GetWindowDesktopId(__RPC__in HWND topLevelWindow, __RPC__out GUID *desktopId) = 0;
  virtual HRESULT STDMETHODCALLTYPE MoveWindowToDesktop(__RPC__in HWND topLevelWindow, __RPC__in REFGUID desktopId) = 0;
};
*/

EXTERN_C const IID IID_IVirtualDesktopNotification;
MIDL_INTERFACE("C179334C-4295-40D3-BEA1-C654D965605A") IVirtualDesktopNotification : public IUnknown{
public:
  virtual HRESULT STDMETHODCALLTYPE VirtualDesktopCreated(IVirtualDesktop *pDesktop) = 0;
  virtual HRESULT STDMETHODCALLTYPE VirtualDesktopDestroyBegin(IVirtualDesktop *pDesktopDestroyed, IVirtualDesktop *pDesktopFallback) = 0;
  virtual HRESULT STDMETHODCALLTYPE VirtualDesktopDestroyFailed(IVirtualDesktop *pDesktopDestroyed, IVirtualDesktop *pDesktopFallback) = 0;
  virtual HRESULT STDMETHODCALLTYPE VirtualDesktopDestroyed(IVirtualDesktop *pDesktopDestroyed, IVirtualDesktop *pDesktopFallback) = 0;
  virtual HRESULT STDMETHODCALLTYPE ViewVirtualDesktopChanged(IApplicationView *pView) = 0;
  virtual HRESULT STDMETHODCALLTYPE CurrentVirtualDesktopChanged(IVirtualDesktop *pDesktopOld, IVirtualDesktop *pDesktopNew) = 0;
};


EXTERN_C const IID IID_IVirtualDesktopNotificationService;
MIDL_INTERFACE("0CD45E71-D927-4F15-8B0A-8FEF525337BF") IVirtualDesktopNotificationService : public IUnknown{
public:
  virtual HRESULT STDMETHODCALLTYPE Register(IVirtualDesktopNotification *pNotification, DWORD *pdwCookie) = 0;
  virtual HRESULT STDMETHODCALLTYPE Unregister(DWORD dwCookie) = 0;
};

EXTERN_C const IID IID_IApplicationViewCollection;
MIDL_INTERFACE("2c08adf0-a386-4b35-9250-0fe183476fcc") IApplicationViewCollection : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient3() = 0;
    virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient4() = 0;
    virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient5() = 0;
    virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient6() = 0;
    virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient7() = 0;
    virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient8() = 0;
    virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient9() = 0;
    virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient10() = 0;
    virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient11() = 0;
    virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient12() = 0;
    virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient13() = 0;
};

static IServiceProvider* init_service_provider() {
  static Microsoft::WRL::ComPtr<IServiceProvider> provider;
  if (provider.Get() == nullptr) {
    auto hr = CoCreateInstance(CLSID_ImmersiveShell, nullptr, CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&provider));
    if (FAILED(hr)) {
      throw std::runtime_error("Cannot open ImmersiveShell provider");
    }
  }
  return provider.Get();
}

static IVirtualDesktopManagerInternal* init_manager_internal() {
  static Microsoft::WRL::ComPtr<IVirtualDesktopManagerInternal> manager;
  if (manager.Get() == nullptr) {
    auto provider = init_service_provider();
    auto hr = provider->QueryService(CLSID_VirtualDesktopAPI_Unknown, IID_PPV_ARGS(&manager));
    if (FAILED(hr)) {
      throw std::runtime_error("Cannot open VirtualDesktop internal manager");
    }
  }
  return manager.Get();
}

static IVirtualDesktopManagerInternal* get_manager_internal() {
  init_com();
  init_service_provider();
  return init_manager_internal();
}

static IVirtualDesktopManager* init_manager() {
  static Microsoft::WRL::ComPtr<IVirtualDesktopManager> manager;
  if (manager.Get() == nullptr) {
    auto provider = init_service_provider();
    auto hr = provider->QueryService(__uuidof(IVirtualDesktopManager), manager.GetAddressOf());
    if (FAILED(hr)) {
      throw std::runtime_error("Cannot open VirtualDesktop manager");
    }
  }
  return manager.Get();
}

static IVirtualDesktopManager* get_manager() {
  init_com();
  init_service_provider();
  return init_manager();
}


void move_window_to_new_desktop(HWND hwnd) {
  auto manager_internal = get_manager_internal();
  auto manager = get_manager();
  Microsoft::WRL::ComPtr<IVirtualDesktop> new_desktop;
  if (FAILED(manager_internal->CreateDesktopW(new_desktop.GetAddressOf())))
    throw std::runtime_error("Cannot create new desktop");
  GUID id;
  if (FAILED(new_desktop->GetID(&id)))
    throw std::runtime_error("Cannot get GUID of the new desktop");
  auto result = manager->MoveWindowToDesktop(hwnd, id);
  /*if (FAILED())
    throw std::runtime_error("Cannot move window");*/
}
