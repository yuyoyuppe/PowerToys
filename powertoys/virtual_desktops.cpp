#include "pch.h"
#include "virtual_desktops.h"

const CLSID CLSID_ImmersiveShell = { 0xC2F03A33, 0x21F5, 0x47FA, 0xB4, 0xBB, 0x15, 0x63, 0x62, 0xA2, 0xF2, 0x39 };
const CLSID CLSID_VirtualDesktopAPI_Unknown = { 0xC5E0CDCA, 0x7B6E, 0x41B2, 0x9F, 0xC4, 0xD9, 0x39, 0x75, 0xCC, 0x46, 0x7B };
const CLSID CLSID_IVirtualDesktopManagerInternal = { 0xF31574D6, 0xB682, 0x4CDC, 0xBD, 0x56, 0x18, 0x27, 0x86, 0x0A, 0xBE, 0xC6 };
const CLSID CLSID_IApplicationViewCollection = { 0x1841C6D7, 0x4F9D, 0x42C0, 0xAF, 0x41, 0x87, 0x47, 0x53, 0x8F, 0x10, 0xE5 };
const CLSID CLSID_IVirtualNotificationService = { 0xA501FDEC, 0x4A09, 0x464C, 0xAE, 0x4E, 0x1B, 0x9C, 0x21, 0xB8, 0x49, 0x18 };

EXTERN_C const IID IID_IApplicationView;
MIDL_INTERFACE("372E1D3B-38D3-42E4-A15B-8AB2B178F513") IApplicationView : public IUnknown{
public:
  virtual HRESULT STDMETHODCALLTYPE SetFocus() = 0;
  virtual HRESULT STDMETHODCALLTYPE SwitchTo() = 0;
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient003() = 0; // TryInvokeBack
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient004() = 0; // GetThumbnailWindow
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient005() = 0; // GetMonitor
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient006() = 0; // GetVisibility
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient007() = 0; // SetCloak
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient008() = 0; // GetPosition
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient009() = 0; // SetPosition
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient010() = 0; // InsertAfterWindow
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient011() = 0; // GetExtendedFramePosition
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient012() = 0; // GetAppUserModelId
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient013() = 0; // SetAppUserModelId
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient014() = 0; // IsEqualByAppUserModelId
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient015() = 0; // GetViewState
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient016() = 0; // SetViewState
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient017() = 0; // GetNeediness
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient018() = 0; // GetLastActivationTimestamp
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient019() = 0; // SetLastActivationTimestamp
  virtual HRESULT STDMETHODCALLTYPE GetVirtualDesktopId(GUID **desktopId) = 0;
  virtual HRESULT STDMETHODCALLTYPE SetVirtualDesktopId(GUID *desktopId) = 0;
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient022() = 0; // GetShowInSwitchers
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient023() = 0; // SetShowInSwitchers
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient024() = 0; // GetScaleFactor
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient025() = 0; // CanReceiveInput
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient026() = 0; // GetCompatibilityPolicyType
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient027() = 0; // SetCompatibilityPolicyType
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient028() = 0; // GetSizeConstraints
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient029() = 0; // GetSizeConstraintsForDpi
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient030() = 0; // SetSizeConstraintsForDpi
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient031() = 0; // OnMinSizePreferencesUpdated
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient032() = 0; // ApplyOperation
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient033() = 0; // IsTray
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient034() = 0; // IsInHighZOrderBand
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient035() = 0; // IsSplashScreenPresented
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient036() = 0; // Flash
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient037() = 0; // GetRootSwitchableOwner
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient038() = 0; // EnumerateOwnershipTree
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient039() = 0; // GetEnterpriseId
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient040() = 0; // IsMirrored
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient041() = 0; // Unknown1
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient042() = 0; // Unknown2
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient043() = 0; // Unknown3
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient044() = 0; // Unknown4
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient045() = 0; // Unknown9
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient046() = 0; // Unknown10
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient047() = 0; // Unknown5
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient048() = 0; // Unknown6
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient049() = 0; // Unknown7
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient050() = 0; // Unknown8
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient051() = 0; // Unknown11
  virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient052() = 0; // Unknown12
};

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
MIDL_INTERFACE("1841C6D7-4F9D-42C0-AF41-8747538F10E5") IApplicationViewCollection : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient001() = 0; // GetViews
    virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient002() = 0; // GetViewsByZOrder
    virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient003() = 0; // GetViewsByAppUserModelId
    virtual HRESULT STDMETHODCALLTYPE GetViewForHwnd(__RPC__in HWND hwnd, IApplicationView** view) = 0;
    virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient005() = 0; // GetViewForApplication
    virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient006() = 0; // GetViewForAppUserModelId
    virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient007() = 0; // GetViewInFocus
    virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient008() = 0; // RefreshCollection
    virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient009() = 0; // RegisterForApplicationViewChanges
    virtual HRESULT STDMETHODCALLTYPE _ObjectStublessClient010() = 0; // UnregisterForApplicationViewChanges
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

  IApplicationViewCollection* get_application_view_collection() {
    auto provider = get_service_provider();
    static winrt::com_ptr<IApplicationViewCollection> collection;
    if(!collection) {
      winrt::check_hresult(provider->QueryService(CLSID_IApplicationViewCollection, __uuidof(collection), collection.put_void()));
    }
    return collection.get();
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
  for (int i = 0; i < (int)count; i++)
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
  if(manager->GetWindowDesktopId(hwnd, &current_desktopId) == S_OK) {
    return GetDesktopGUIDIndex(current_desktopId);
  } else {
    // Unable to get the Desktop ID for the window.
    return -1;
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

// TODO: Add a synchronized map instead, after testing this technique.
std::map<HWND, WINDOWPLACEMENT> moved_window_original_positions;

BOOL CALLBACK check_if_window_in_virtual_desktop(HWND hwnd, LPARAM ptrGUID)  {
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

void switch_to_primary_desktop_and_delete_after_delay(HWND hwnd, GUID old_desktop_id, GUID primary_desktop_id, int ms) {
  if (ms>0) {
    std::this_thread::sleep_for(std::chrono::milliseconds(ms));
  }

  BOOL move_success = FALSE;
  auto current_foreground_window = GetForegroundWindow();
  if ( current_foreground_window != hwnd) {
    move_success=SetForegroundWindow(hwnd);
  }

  auto manager_internal = get_manager_internal();
  winrt::com_ptr<IVirtualDesktop> primary_desktop = nullptr;
  winrt::com_ptr<IVirtualDesktop> old_desktop = nullptr;
  if (manager_internal->FindDesktop(&primary_desktop_id, primary_desktop.put()) == S_OK &&
      manager_internal->FindDesktop(&old_desktop_id, old_desktop.put()) == S_OK) {
    if(!move_success) {
      // Couldn't set hwnd as the ForegroundWindow or it was the same window.
      // Will have to switch desktop manually, without animation.
      winrt::check_hresult(manager_internal->SwitchDesktop(primary_desktop.get()));
    } else {
      // Give time for the animation to play before trying to remove the old desktop.
      std::this_thread::sleep_for(std::chrono::milliseconds(ms));
    }
    if (EnumWindows(check_if_window_in_virtual_desktop, reinterpret_cast<LPARAM>(&old_desktop_id)) != FALSE) {
      manager_internal->RemoveDesktop(old_desktop.get(), primary_desktop.get());
    } else {
      if(move_success) {
        // Making sure we do switch desktops, in case the animation doesn't play.
        winrt::check_hresult(manager_internal->SwitchDesktop(primary_desktop.get()));
      }
    }
  }
}

void move_window_to_primary_desktop(HWND hwnd) {
  auto manager_internal = get_manager_internal();
  auto manager = get_manager();

  GUID current_desktopId;
  winrt::check_hresult(manager->GetWindowDesktopId(hwnd, &current_desktopId));

  // Restore the window to the original position.
  if (moved_window_original_positions.find(hwnd) != moved_window_original_positions.end()) {
    WINDOWPLACEMENT original_placement = moved_window_original_positions[hwnd];
  SetWindowPlacement(hwnd, &original_placement);
  moved_window_original_positions.erase(hwnd);
  }

  int desktop_index = 0;
  winrt::com_ptr<IVirtualDesktop> objDestkop;
  try {
    objDestkop = GetDesktopAtIndex(desktop_index);
  }
  catch (std::out_of_range) {
    // Tried to search the GUID of an out of range index desktop.
    MessageBox(NULL, "Primary desktop index is out of range.", "Error", MB_OK | MB_ICONERROR);
    return;
  }
  GUID primary_desktopId;
  winrt::check_hresult(objDestkop->GetID(&primary_desktopId));

  auto collection_view = get_application_view_collection();

  winrt::com_ptr <IApplicationView> view;
  winrt::check_hresult(collection_view->GetViewForHwnd(hwnd, view.put()));
  winrt::check_hresult(manager_internal->MoveViewToDesktop(view.get(), objDestkop.get()));

  auto current_foreground_window = GetForegroundWindow();
  if ( current_foreground_window != hwnd) {
    // Switching the moved window to the foreground will animate the transition.
    std::thread(switch_to_primary_desktop_and_delete_after_delay, hwnd, current_desktopId, primary_desktopId, 50).detach();
  } else {
    auto tasklist_hwnd = FindWindow("Shell_TrayWnd", nullptr);
    if (!SetForegroundWindow(tasklist_hwnd)) {
      // Couldn't set TaskBar as the foreground Window. Will just move to new Desktop regardless.
      std::thread(switch_to_primary_desktop_and_delete_after_delay, hwnd, current_desktopId, primary_desktopId, 50).detach();
    } else {
      // Try to switch to new Desktop after a bit.
      std::thread(switch_to_primary_desktop_and_delete_after_delay, hwnd, current_desktopId, primary_desktopId, 50).detach();
    }
  }

}

void switch_to_window_desktop_after_delay(HWND hwnd, GUID new_desktop_id, int ms) {
  if (ms>0) {
    std::this_thread::sleep_for(std::chrono::milliseconds(ms));
  }
  BOOL move_success = FALSE;
  auto current_foreground_window = GetForegroundWindow();
  if ( current_foreground_window != hwnd) {
    move_success=SetForegroundWindow(hwnd);
  }

  // In the case we couldn't set hwnd as the ForegroundWindow or it was the same window.
  // Will have to switch desktop manually, without animation.
  std::this_thread::sleep_for(std::chrono::milliseconds(ms));

  auto manager_internal = get_manager_internal();
  winrt::com_ptr<IVirtualDesktop> new_desktop;
  if (manager_internal->FindDesktop(&new_desktop_id, new_desktop.put()) == S_OK) {
    winrt::check_hresult(manager_internal->SwitchDesktop(new_desktop.get()));
  }

};

void move_window_to_new_desktop(HWND hwnd) {
  auto manager_internal = get_manager_internal();
  auto manager = get_manager();
  winrt::com_ptr<IVirtualDesktop> new_desktop;
  winrt::check_hresult(manager_internal->CreateDesktopW(new_desktop.put()));
  GUID id;
  winrt::check_hresult(new_desktop->GetID(&id));

  auto collection_view = get_application_view_collection();

  WINDOWPLACEMENT original_placement;
  original_placement.length = sizeof(WINDOWPLACEMENT);
  if (GetWindowPlacement(hwnd, &original_placement)) {
    moved_window_original_positions[hwnd] = original_placement;
  }

  ShowWindow(hwnd, SW_MAXIMIZE); // Maximize before moving.
  
  winrt::com_ptr <IApplicationView> view;
  winrt::check_hresult(collection_view->GetViewForHwnd(hwnd, view.put()));
  winrt::check_hresult(manager_internal->MoveViewToDesktop(view.get(), new_desktop.get()));

  auto current_foreground_window = GetForegroundWindow();
  if ( current_foreground_window != hwnd) {
    // Switching the moved window to the foreground will animate the transition.
    switch_to_window_desktop_after_delay(hwnd, id, 50);
  } else {
    auto tasklist_hwnd = FindWindow("Shell_TrayWnd", nullptr);
    if (!SetForegroundWindow(tasklist_hwnd)) {
      // Couldn't set TaskBar as the foreground Window. Will just move to new Desktop regardless.
      switch_to_window_desktop_after_delay(hwnd, id, 50);
    } else {
      // Try to switch to new Desktop after a bit.
      std::thread(switch_to_window_desktop_after_delay, hwnd, id, 50).detach();
    }
  }
}
