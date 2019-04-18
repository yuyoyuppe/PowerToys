#include "pch.h"
#include "virtual_desktops.h"
#include "move_window.h"

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
int GetDesktopGUIDIndex(GUID id) {
  auto manager_internal = get_manager_internal();
  int index = -1;
  u_int count;
  winrt::check_hresult(manager_internal->GetCount(&count));
  winrt::com_ptr<IObjectArray> desktops;
  manager_internal->GetDesktops(desktops.put());
  for (int i = 0; i < count; i++)
  {
    winrt::com_ptr<IVirtualDesktop> objdesktop;
    desktops.get()->GetAt(i, __uuidof(IVirtualDesktop), objdesktop.put_void());
    GUID compare_id;
    winrt::check_hresult(objdesktop->GetID(&compare_id));
    if (IsEqualGUID(id, compare_id)) {
      index = i;
      break;
    }
  }
  //TODO: Verify releases needed with IObjectArray and winrt::com_ptr
  return index;
}

int GetCurrentDesktopGUIDIndexForWindow(HWND hwnd) {
  auto manager = get_manager();
  GUID current_desktopId;
  winrt::check_hresult(manager->GetWindowDesktopId(hwnd, &current_desktopId));
  return GetDesktopGUIDIndex(current_desktopId);
}

// Adapted from https://gallery.technet.microsoft.com/scriptcenter/Powershell-commands-to-d0e79cc5
winrt::com_ptr<IVirtualDesktop> GetDesktopAtIndex(int index) {
  auto manager_internal = get_manager_internal();
  UINT count;
  winrt::check_hresult(manager_internal->GetCount(&count));
  if (index < 0 || index >= count) throw std::out_of_range("GetDesktopGUIDAtIndex : index is out of range.");
  winrt::com_ptr<IObjectArray> desktops;
  manager_internal->GetDesktops(desktops.put());
  winrt::com_ptr<IVirtualDesktop> objdesktop;
  desktops.get()->GetAt(index, __uuidof(IVirtualDesktop), objdesktop.put_void());
  //TODO: Verify releases needed with IObjectArray and winrt::com_ptr
  return objdesktop;
}

#define MOVETONEWDESKTOPMSGSTR "POWERTOYS_MOVE_TO_NEW_DESKTOP"

void move_window_to_primary_desktop(HWND hwnd) {
  auto manager_internal = get_manager_internal();
  auto manager = get_manager();
  // Restore the Window.
  ShowWindow(hwnd, SW_RESTORE);
  UINT msg = RegisterWindowMessage(MOVETONEWDESKTOPMSGSTR);

  int desktop_index = 0;
  winrt::com_ptr<IVirtualDesktop> objDestkop;
  try {
    objDestkop = GetDesktopAtIndex(desktop_index);
  }
  catch (std::out_of_range& ex) {
    // Tried to search the GUID of an out of range index desktop.
    MessageBox(NULL, "Primary desktop index is out of range.", "Error", MB_OK | MB_ICONERROR);
    return;
  }

  // Send custom message, with target desktop index in LPARAM.
  BOOL res = PostMessage(hwnd, msg, 0, (LPARAM)desktop_index);
  if (res==0) {
    DWORD dw = GetLastError();
    if (dw == 5) {
      // Access denied. Means Windows UIPI is blocking powertoys from moving the window to another Desktop.
      MessageBox(NULL, "Couldn't move the window to a new Desktop. Need to start as an Administrator to do that.", "Access Denied", MB_OK | MB_ICONEXCLAMATION);
    }
    else {
      ShowLastErrorMessage((LPTSTR)"PostMessage", dw);
    }
    return;
  }

  // Switch to the new desktop.
  winrt::check_hresult(manager_internal->SwitchDesktop(objDestkop.get()));
}

void move_window_to_new_desktop(HWND hwnd) {
  auto manager_internal = get_manager_internal();
  auto manager = get_manager();
  winrt::com_ptr<IVirtualDesktop> new_desktop;
  winrt::check_hresult(manager_internal->CreateDesktopW(new_desktop.put()));
  GUID id;
  winrt::check_hresult(new_desktop->GetID(&id));
  
  // Maximize the Window.
  ShowWindow(hwnd, SW_MAXIMIZE);

  // This fails, because it needs to run in the target Window's process context.
  // manager->MoveWindowToDesktop(hwnd, id);
  // Instead, we use executables to install a message hook from 64 and 32 bits dlls
  // at startup and Post a Message from here.

  UINT msg = RegisterWindowMessage(MOVETONEWDESKTOPMSGSTR);
  int desktop_index = GetDesktopGUIDIndex(id);

  // Send custom message, with target desktop index in LPARAM.
  BOOL res = PostMessage(hwnd, msg, 0, (LPARAM)desktop_index);
  if (res==0) {
    DWORD dw = GetLastError();
    // Couldn't ask the Window to Move. Remove the created Desktop.
    winrt::com_ptr<IVirtualDesktop> curr_desktop;
    winrt::check_hresult(manager_internal->GetCurrentDesktop(curr_desktop.put()));
    winrt::check_hresult(manager_internal->RemoveDesktop(new_desktop.get(), curr_desktop.get()));
    if (dw == 5) {
      // Access denied. Means Windows UIPI is blocking powertoys from moving the window to another Desktop.
      MessageBox(NULL, "Couldn't move the window to a new Desktop. Need to start as an Administrator to do that.", "Access Denied", MB_OK | MB_ICONEXCLAMATION);
    }
    else {
      ShowLastErrorMessage((LPTSTR)"PostMessage", dw);
    }
    return;
  }

  // Switch to the new desktop.
  winrt::check_hresult(manager_internal->SwitchDesktop(new_desktop.get()));
}
