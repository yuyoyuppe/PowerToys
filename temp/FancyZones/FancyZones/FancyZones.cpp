#include "stdafx.h"

using namespace Microsoft::WRL;

void HookMoveSizeStart(_In_ HWND window, POINT const& ptScreen) noexcept;
void HookMoveSizeEnd(_In_ HWND window, POINT const& ptScreen) noexcept;
void HookMoveSizeUpdate(POINT const& ptScreen) noexcept;

class FancyZones WrlFinal : public RuntimeClass<
    RuntimeClassFlags<ClassicCom>,
    IFancyZones>
{
public:
    FancyZones(_In_ HINSTANCE hinstance) noexcept :
        m_hinstance(hinstance)
    {
    }

    IFACEMETHODIMP_(HWND) GetWindow() noexcept { return m_window; }
    IFACEMETHODIMP_(HINSTANCE) GetHInstance() noexcept { return m_hinstance; }
    IFACEMETHODIMP_(Settings) GetSettings() noexcept { return m_settings->GetSettings(); }
    IFACEMETHODIMP_(void) Run() noexcept;
    IFACEMETHODIMP_(void) ToggleZoneViewers() noexcept;
    IFACEMETHODIMP_(void) ShowZoneEditorForMonitor(_In_ HMONITOR monitor) noexcept;
    IFACEMETHODIMP_(bool) InMoveSize() noexcept { return m_inMoveSize; }
    IFACEMETHODIMP MoveSizeEnter(_In_ HWND window, _In_ HMONITOR monitor, POINT ptScreen) noexcept;
    IFACEMETHODIMP MoveSizeExit(_In_ HWND window, POINT ptScreen) noexcept;
    IFACEMETHODIMP MoveSizeUpdate(_In_ HMONITOR monitor, POINT ptScreen) noexcept;
    IFACEMETHODIMP AddZoneWindow(_In_ IZoneWindow* zoneWindow, _In_ HMONITOR monitor) noexcept;
    IFACEMETHODIMP_(void) OnDisplayChange(DisplayChangeType changeType) noexcept;
    IFACEMETHODIMP_(void) MoveWindowIntoZoneByIndex(_In_ HWND window, int index) noexcept;
    IFACEMETHODIMP_(bool) IsInterestingWindow(_In_ HWND window) noexcept;
    IFACEMETHODIMP_(void) VirtualDesktopChanged() noexcept;
    IFACEMETHODIMP_(GUID) GetCurrentVirtualDesktopId() noexcept { return m_currentVirtualDesktopId; }
    IFACEMETHODIMP_(bool) OnKeyDown(LPARAM lparam) noexcept;
    IFACEMETHODIMP_(void) MoveWindowsOnActiveZoneSetChange() noexcept;

protected:
    ~FancyZones() noexcept;
    static LRESULT CALLBACK s_WndProc(HWND, UINT, WPARAM, LPARAM) noexcept;
    static void CALLBACK s_HookProc(_In_ HWINEVENTHOOK winEventHook, DWORD event, _In_ HWND window, long object, long child, DWORD eventThread, DWORD eventTime) noexcept;
    static LRESULT CALLBACK s_keyboardHookProc(int code, WPARAM wparam, LPARAM lparam) noexcept;

private:
    void InitialAddZoneWindows() noexcept;
    void MoveWindowsOnDisplayChange() noexcept;
    void UpdateDragMode() noexcept;
    void CycleActiveZoneSet(DWORD vkCode) noexcept;
    void OnSnapHotkey(DWORD vkCode) noexcept;

    HINSTANCE m_hinstance{};
    HWINEVENTHOOK m_hook{};
    HHOOK m_keyboardHook{};
    HWND m_window{};
    HWND m_windowMoveSize{};
    bool m_editorsVisible{};
    bool m_inMoveSize{};
    std::map<HMONITOR, ComPtr<IZoneWindow>> m_zoneWindowMap;
    ComPtr<IZoneWindow> m_zoneWindowMoveSize;
    DragMode m_dragMode = DragMode::None;
    ComPtr<IFancyZonesSettings> m_settings;
    GUID m_currentVirtualDesktopId{};
};

FancyZones::~FancyZones() noexcept
{
    if (m_window)
    {
        DestroyWindow(m_window);
        m_window = nullptr;
    }
}

IFACEMETHODIMP_(void) FancyZones::Run() noexcept
{
    WNDCLASSEXW wcex{};
    wcex.cbSize = sizeof(WNDCLASSEX);
    wcex.lpfnWndProc = s_WndProc;
    wcex.hInstance = m_hinstance;
    wcex.lpszClassName = L"SuperFancyZones";
    RegisterClassExW(&wcex);

    m_hook = SetWinEventHook(EVENT_SYSTEM_MOVESIZESTART, EVENT_OBJECT_NAMECHANGE, nullptr, s_HookProc, 0, 0, WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS);
    if (!m_hook) return;

    m_keyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, s_keyboardHookProc, nullptr, 0);
    if (!m_keyboardHook) return;

    BufferedPaintInit();

    m_window = CreateWindowExW(WS_EX_TOOLWINDOW, L"SuperFancyZones", L"", WS_POPUP, 0, 0, 0, 0, nullptr, nullptr, m_hinstance, nullptr);
    if (!m_window) return;

    m_settings = MakeFancyZonesSettings();

    RegisterHotKey(m_window, 1, MOD_WIN, VK_OEM_3);
    VirtualDesktopChanged();

    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    BufferedPaintUnInit();
    UnhookWindowsHookEx(m_keyboardHook);
    UnhookWinEvent(m_hook);
}

IFACEMETHODIMP_(void) FancyZones::ToggleZoneViewers() noexcept
{
    if (!m_editorsVisible)
    {
        auto callback = [](_In_ HMONITOR monitor, _In_ HDC, _In_ RECT *, _In_ LPARAM dwData) -> BOOL
        {
            g_app->ShowZoneEditorForMonitor(monitor);
            return TRUE;
        };
        EnumDisplayMonitors(nullptr, nullptr, callback, 0);
    }
    else
    {
        for (auto iter : m_zoneWindowMap)
        {
            iter.second->HideZoneWindow();
        }
    }
    m_editorsVisible = !m_editorsVisible;
}

IFACEMETHODIMP_(void) FancyZones::ShowZoneEditorForMonitor(_In_ HMONITOR monitor) noexcept
{
    auto iter = m_zoneWindowMap.find(monitor);
    if (iter != m_zoneWindowMap.end())
    {
        bool const activate = MonitorFromPoint(POINT(), MONITOR_DEFAULTTOPRIMARY) == monitor;
        iter->second->ShowZoneWindow(activate);

    }
}

IFACEMETHODIMP FancyZones::MoveSizeEnter(_In_ HWND window, _In_ HMONITOR monitor, POINT ptScreen) noexcept
{
    // Only enter move/size if the cursor is in the titlebar.
    // This prevents resize from triggering zones.
    RECT windowRect{};
    ::GetWindowRect(window, &windowRect);

    TITLEBARINFO titlebarInfo{sizeof(titlebarInfo)};
    ::GetTitleBarInfo(window, &titlebarInfo);

    // Titlebar height is weird and apps can do custom drag areas.
    // Give it most of the height of the window to make sure.
    titlebarInfo.rcTitleBar.bottom += ((windowRect.bottom - windowRect.top) / 2);

    if (PtInRect(&titlebarInfo.rcTitleBar, ptScreen))
    {
        m_inMoveSize = true;
        auto iter = m_zoneWindowMap.find(monitor);
        if (iter != m_zoneWindowMap.end())
        {
            m_windowMoveSize = window;

            UpdateDragMode();
            if (m_dragMode != DragMode::None)
            {
                m_zoneWindowMoveSize = iter->second;
                return m_zoneWindowMoveSize->MoveSizeEnter(window, ptScreen, m_dragMode);
            }
            else if (m_zoneWindowMoveSize)
            {
                m_zoneWindowMoveSize->MoveSizeCancel();
                m_zoneWindowMoveSize = nullptr;
            }
            return E_FAIL;
        }
    }
    return E_INVALIDARG;
}

IFACEMETHODIMP FancyZones::MoveSizeExit(_In_ HWND window, POINT ptScreen) noexcept
{
    m_inMoveSize = false;
    m_windowMoveSize = nullptr;
    m_dragMode = DragMode::None;
    if (m_zoneWindowMoveSize)
    {
        auto zoneWindow = std::move(m_zoneWindowMoveSize);
        return zoneWindow->MoveSizeExit(window, ptScreen);
    }
    else
    {
        ::RemoveProp(window, ZONE_STAMP);
    }
    return E_INVALIDARG;
}

IFACEMETHODIMP FancyZones::MoveSizeUpdate(_In_ HMONITOR monitor, POINT ptScreen) noexcept
{
    if (m_inMoveSize)
    {
        UpdateDragMode();
        
        if (m_zoneWindowMoveSize)
        {
            if (m_dragMode == DragMode::None)
            {
                auto zoneWindow = std::move(m_zoneWindowMoveSize);
                return zoneWindow->MoveSizeCancel();
            }
            else
            {
                auto iter = m_zoneWindowMap.find(monitor);
                if (iter != m_zoneWindowMap.end())
                {
                    if (iter->second != m_zoneWindowMoveSize)
                    {
                        auto const dragMode = m_zoneWindowMoveSize->GetDragMode();
                        m_zoneWindowMoveSize->MoveSizeCancel();
                        m_zoneWindowMoveSize = iter->second;
                        m_zoneWindowMoveSize->MoveSizeEnter(m_windowMoveSize, ptScreen, dragMode);
                    }
                    return m_zoneWindowMoveSize->MoveSizeUpdate(ptScreen, m_dragMode);
                }
                return E_INVALIDARG;
            }
        }
        else if (m_dragMode != DragMode::None)
        {
            HookMoveSizeStart(m_windowMoveSize, ptScreen);
            HookMoveSizeUpdate(ptScreen);
            return S_OK;
        }
    }
    return E_INVALIDARG;
}

IFACEMETHODIMP FancyZones::AddZoneWindow(_In_ IZoneWindow* zoneWindow, _In_ HMONITOR monitor) noexcept
{
    auto iter = m_zoneWindowMap.find(monitor);
    if (iter != m_zoneWindowMap.end())
    {
        // XXXX: this shouldn't happen
        m_zoneWindowMap.erase(iter);
    }

    m_zoneWindowMap[monitor] = zoneWindow;
    return S_OK;
}

IFACEMETHODIMP_(void) FancyZones::OnDisplayChange(DisplayChangeType changeType) noexcept
{
    // Do this on a delay so we don't do too much work?
    // Determine if any resolutions changed
    // Should probably support DPI...
    // Update all zone windows
    // Delete invalid ones
    // Create new ones
    // Keep windows in zones

    for (auto iter : m_zoneWindowMap)
    {
        // XXXX: OnDisplayChange makes an outoing com call that pumps messages
        // So we can reenter if we get multiple messages quickly
        // Need to protect against that
        iter.second->OnDisplayChange(changeType);
    }

    m_zoneWindowMap.clear();
    InitialAddZoneWindows();

    if (((changeType == DisplayChangeType::WorkArea) ||
         (changeType == DisplayChangeType::DisplayChange) &&
         m_settings->GetSettings().displayChange_moveWindows) ||
        ((changeType == DisplayChangeType::VirtualDesktop) &&
         m_settings->GetSettings().virtualDesktopChange_moveWindows))
    {
        MoveWindowsOnDisplayChange();
    }
}

IFACEMETHODIMP_(void) FancyZones::MoveWindowIntoZoneByIndex(_In_ HWND window, int index) noexcept
{
    if (window != m_windowMoveSize)
    {
        HMONITOR monitor = MonitorFromWindow(window, MONITOR_DEFAULTTONULL);
        if (monitor)
        {
            auto iter = m_zoneWindowMap.find(monitor);
            if (iter != m_zoneWindowMap.end())
            {
                iter->second->MoveWindowIntoZoneByIndex(window, index);
            }
        }
    }
}

IFACEMETHODIMP_(bool) FancyZones::IsInterestingWindow(_In_ HWND window) noexcept
{
    return IsFlagSet(GetWindowLongPtr(window, GWL_STYLE), WS_MAXIMIZEBOX);
}

IFACEMETHODIMP_(void) FancyZones::VirtualDesktopChanged() noexcept
{
    GUID currentVirtualDesktopId{};
    if (SUCCEEDED(RegistryHelpers::GetCurrentVirtualDesktop(&currentVirtualDesktopId)))
    {
        m_currentVirtualDesktopId = currentVirtualDesktopId;
        OnDisplayChange(DisplayChangeType::VirtualDesktop);
    }
}

IFACEMETHODIMP_(bool) FancyZones::OnKeyDown(LPARAM lparam) noexcept
{
    auto info = reinterpret_cast<PKBDLLHOOKSTRUCT>(lparam);
    bool const alt = GetAsyncKeyState(VK_MENU) & 0x8000;
    bool const shift = GetAsyncKeyState(VK_SHIFT) & 0x8000;
    bool const win = GetAsyncKeyState(VK_LWIN) & 0x8000;
    if (win && !alt && !shift)
    {
        if (!m_settings->GetSettings().overrideSnapHotkeys)
        {
            return false;
        }

        bool const ctrl = GetAsyncKeyState(VK_CONTROL) & 0x8000;
        if (ctrl)
        {
            if ((info->vkCode >= '0') && (info->vkCode <= '9'))
            {
                CycleActiveZoneSet(info->vkCode);
                return true;
            }
        }
        else
        {
            // XXXX: better directional control
            // VK_UP/VK_DOWN, use zone rects to figure out most appropriate given direction
            if ((info->vkCode == VK_RIGHT) || (info->vkCode == VK_LEFT))
            {
                OnSnapHotkey(info->vkCode);
                return true;
            }
        }
    }
    else if (m_inMoveSize && (info->vkCode >= '0') && (info->vkCode <= '9'))
    {
        CycleActiveZoneSet(info->vkCode);
        return true;
    }
    return false;
}

IFACEMETHODIMP_(void) FancyZones::MoveWindowsOnActiveZoneSetChange() noexcept
{
    if (m_settings->GetSettings().zoneSetChange_moveWindows)
    {
        MoveWindowsOnDisplayChange();
    }
}

LRESULT CALLBACK FancyZones::s_WndProc(_In_ HWND window, UINT message, _In_ WPARAM wparam, _In_ LPARAM lparam) noexcept
{
    switch (message)
    {
        case WM_DESTROY:
        {
            PostQuitMessage(0);
        }
        break;

        case WM_HOTKEY:
        {
            if (wparam == 1)
            {
                g_app->ToggleZoneViewers();
            }
        }
        break;

        case WM_SETTINGCHANGE:
        {
            if (wparam == SPI_SETWORKAREA)
            {
                g_app->OnDisplayChange(DisplayChangeType::WorkArea);
            }
        }
        break;

        case WM_DISPLAYCHANGE:
        {
            g_app->OnDisplayChange(DisplayChangeType::DisplayChange);
        }
        break;

        default:
        {
            return DefWindowProc(window, message, wparam, lparam);
        }
    }
    return 0;
}

void HookMoveSizeStart(_In_ HWND window, POINT const& ptScreen) noexcept
{
    if (g_app->IsInterestingWindow(window))
    {
        auto monitor = MonitorFromPoint(ptScreen, MONITOR_DEFAULTTONULL);
        if (monitor)
        {
            g_app->MoveSizeEnter(window, monitor, ptScreen);
        }
    }
}

void HookMoveSizeEnd(_In_ HWND window, POINT const& ptScreen) noexcept
{
    if (g_app->IsInterestingWindow(window))
    {
        g_app->MoveSizeExit(window, ptScreen);
    }
}

void HookMoveSizeUpdate(POINT const& ptScreen) noexcept
{
    auto monitor = MonitorFromPoint(ptScreen, MONITOR_DEFAULTTONULL);
    if (monitor)
    {
        g_app->MoveSizeUpdate(monitor, ptScreen);
    }
}

void CALLBACK FancyZones::s_HookProc(
    _In_ HWINEVENTHOOK /*winEventHook*/,
    DWORD event,
    _In_ HWND window,
    long /*object*/,
    long /*child*/,
    DWORD /*eventThread*/,
    DWORD /*eventTime*/) noexcept
{
    if (g_app->GetWindow() != nullptr)
    {
        POINT ptScreen;
        GetPhysicalCursorPos(&ptScreen);

        switch (event)
        {
        case EVENT_SYSTEM_MOVESIZESTART:
            HookMoveSizeStart(window, ptScreen);
            break;

        case EVENT_SYSTEM_MOVESIZEEND:
            HookMoveSizeEnd(window, ptScreen);
            break;

        case EVENT_OBJECT_LOCATIONCHANGE:
            if (g_app->InMoveSize())
            {
                HookMoveSizeUpdate(ptScreen);
            }
            break;

        case EVENT_OBJECT_NAMECHANGE:
            {
                if (window == GetDesktopWindow())
                {
                    g_app->VirtualDesktopChanged();
                }
            }
            break;
        }
    }
}

LRESULT CALLBACK FancyZones::s_keyboardHookProc(int code, WPARAM wparam, LPARAM lparam) noexcept
{
    if ((code >= 0) && (wparam == WM_KEYDOWN) && g_app->OnKeyDown(lparam))
    {
        return 1;
    }
    return CallNextHookEx(nullptr, code, wparam, lparam);
}

void FancyZones::InitialAddZoneWindows() noexcept
{
    auto callback = [](_In_ HMONITOR monitor, _In_ HDC, _In_ RECT *, _In_ LPARAM dwData) -> BOOL
    {
        MONITORINFOEX mi;
        mi.cbSize = sizeof(mi);
        if (GetMonitorInfo(monitor, &mi))
        {
            DISPLAY_DEVICE displayDevice = { sizeof(displayDevice) };
            PCWSTR deviceId = nullptr;

            bool validMonitor = true;
            if (EnumDisplayDevices(mi.szDevice, 0, &displayDevice, 1))
            {
                if (IsFlagSet(displayDevice.StateFlags, DISPLAY_DEVICE_MIRRORING_DRIVER))
                {
                    validMonitor = FALSE;
                }
                else if (displayDevice.DeviceID[0] != L'\0')
                {
                    deviceId = displayDevice.DeviceID;
                }
            }

            if (validMonitor)
            {
                if (!deviceId)
                {
                    deviceId = GetSystemMetrics(SM_REMOTESESSION) ?
                        L"\\\\?\\DISPLAY#REMOTEDISPLAY#" :
                        L"\\\\?\\DISPLAY#LOCALDISPLAY#";
                }

                auto zoneWindow = MakeZoneWindow(monitor, deviceId);
                if (zoneWindow)
                {
                    g_app->AddZoneWindow(zoneWindow.Get(), monitor);
                }
            }
        }
        return TRUE;
    };
    EnumDisplayMonitors(nullptr, nullptr, callback, 0);
}

void FancyZones::MoveWindowsOnDisplayChange() noexcept
{
    auto callback = [](_In_ HWND window, _In_ LPARAM lparam) -> BOOL
    {
        int i = reinterpret_cast<int>(::GetProp(window, ZONE_STAMP));
        if (i != 0)
        {
            // i is off by 1 since 0 is special.
            g_app->MoveWindowIntoZoneByIndex(window, i-1);
        }
        return TRUE;
    };
    EnumWindows(callback, 0);
}

void FancyZones::UpdateDragMode() noexcept
{
    bool const alt = GetAsyncKeyState(VK_MENU) & 0x8000;
    bool const shift = GetAsyncKeyState(VK_SHIFT) & 0x8000;

    m_dragMode = m_settings->GetSettings().dragMode;
    if (alt)
    {
        m_dragMode = (m_dragMode == DragMode::SetWindowPos) ? DragMode::None : DragMode::SetWindowPos;
    }
    else if (shift)
    {
        m_dragMode = (m_dragMode == DragMode::Normal) ? DragMode::None : DragMode::Normal;
    }
}

void FancyZones::CycleActiveZoneSet(DWORD vkCode) noexcept
{
    HWND const window = GetForegroundWindow();
    if (window != nullptr)
    {
        HMONITOR monitor = MonitorFromWindow(window, MONITOR_DEFAULTTONULL);
        if (monitor)
        {
            auto iter = m_zoneWindowMap.find(monitor);
            if (iter != m_zoneWindowMap.end())
            {
                iter->second->CycleActiveZoneSet(vkCode);
            }
        }
    }
}

void FancyZones::OnSnapHotkey(DWORD vkCode) noexcept
{
    HWND const window = GetForegroundWindow();
    if (window != nullptr)
    {
        HMONITOR monitor = MonitorFromWindow(window, MONITOR_DEFAULTTONULL);
        if (monitor)
        {
            auto iter = m_zoneWindowMap.find(monitor);
            if (iter != m_zoneWindowMap.end())
            {
                iter->second->MoveWindowIntoZoneByDirection(window, vkCode);
            }
        }
    }
}

ComPtr<IFancyZones> MakeFancyZones(_In_ HINSTANCE hinstance) noexcept
{
    return Make<FancyZones>(hinstance);
}
