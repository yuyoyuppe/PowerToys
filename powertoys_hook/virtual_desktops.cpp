#include "virtual_desktops.h"
#include "stdafx.h"
#include <winrt/base.h>
#include <Windows.h>
#include <Shobjidl.h>

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

namespace {
  IServiceProvider* get_service_provider() {
    static winrt::com_ptr<IServiceProvider> provider;
    if (!provider) {
      winrt::check_hresult(CoCreateInstance(CLSID_ImmersiveShell, nullptr, CLSCTX_LOCAL_SERVER, __uuidof(provider), provider.put_void()));
    }
    return provider.get();
  }

  IVirtualDesktopManagerInternal* get_manager_internal() {
    auto provider = get_service_provider();
    static winrt::com_ptr<IVirtualDesktopManagerInternal> manager;
    if (!manager) {
      winrt::check_hresult(provider->QueryService(CLSID_VirtualDesktopAPI_Unknown, __uuidof(manager), manager.put_void()));
    }
    return manager.get();
  }

  IVirtualDesktopManager* get_manager() {
    auto provider = get_service_provider();
    static winrt::com_ptr<IVirtualDesktopManager> manager;
    if (!manager) {
      winrt::check_hresult(provider->QueryService(__uuidof(manager), manager.put()));
    }
    return manager.get();
  }
}

// Adapted from https://gallery.technet.microsoft.com/scriptcenter/Powershell-commands-to-d0e79cc5
winrt::com_ptr<IVirtualDesktop> GetDesktopAtIndex(int index) {
  auto manager_internal = get_manager_internal();
  UINT count;
  winrt::check_hresult(manager_internal->GetCount(&count));
  if (index < 0 || index >= (int)count) throw std::out_of_range("GetDesktopGUIDAtIndex : index is out of range.");
  winrt::com_ptr<IObjectArray> desktops;
  manager_internal->GetDesktops(desktops.put());
  winrt::com_ptr<IVirtualDesktop> objdesktop;
  desktops.get()->GetAt(index, __uuidof(IVirtualDesktop), objdesktop.put_void());
  //TODO: Verify releases needed with IObjectArray and winrt::com_ptr
  return objdesktop;
}

BOOL CALLBACK CheckIfWindowInVirtualDesktop(HWND hwnd, LPARAM ptrGUID)  {
  if (!hwnd) {
    return TRUE;
  }
  auto manager = get_manager();
  GUID test_desktopId;
  if(manager->GetWindowDesktopId(hwnd, &test_desktopId) != S_OK) {
    // Couldn't get the DesktopId for the Window.
    return TRUE;
  }
  if( test_desktopId == *(reinterpret_cast<GUID*>(ptrGUID)) ) {
    // This Window is in the desktop we're checking against.
    return FALSE;
  }
  return TRUE;
}

void move_window_to_new_desktop_impl(HWND hwnd, int desktop_index) {
  auto manager_internal = get_manager_internal();
  auto manager = get_manager();

  GUID current_desktopId;
  winrt::check_hresult(manager->GetWindowDesktopId(hwnd, &current_desktopId));

  winrt::com_ptr<IVirtualDesktop> target_desktop;
  GUID target_desktopId;
  try {
    target_desktop = GetDesktopAtIndex(desktop_index);
    winrt::check_hresult(target_desktop->GetID(&target_desktopId));
  }
  catch (std::out_of_range) {
    // Tried to search the GUID of an out of range index desktop.
    MessageBox(NULL, "Desktop index is out of range.", "Error", MB_OK | MB_ICONERROR);
    return;
  }
  if (IsEqualGUID(current_desktopId, target_desktopId)) {
    // Duplicate move request to the same desktop.
    // Ignore it.
    return;
  }

  // Move Window To Desktop
  winrt::check_hresult(manager->MoveWindowToDesktop(hwnd, target_desktopId));
  // Switch to the new desktop.
  winrt::check_hresult(manager_internal->SwitchDesktop(target_desktop.get()));

  if (desktop_index == 0) {
    // Returning the Window to Primary Desktop. Verify if we can delete it.
    winrt::com_ptr<IVirtualDesktop> current_desktop;
    winrt::check_hresult(manager_internal->FindDesktop(&current_desktopId, current_desktop.put()));
    if (EnumWindows(CheckIfWindowInVirtualDesktop, reinterpret_cast<LPARAM>(&current_desktopId)) != FALSE) {
      winrt::check_hresult(manager_internal->RemoveDesktop(current_desktop.get(), target_desktop.get()));
    }
  }

  // Set The focus to the switched window.
  SetForegroundWindow(hwnd);
}
