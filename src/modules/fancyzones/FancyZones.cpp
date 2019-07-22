#include "pch.h"

struct FancyZones : public winrt::implements<FancyZones, IFancyZones, IFancyZonesCallback, IZoneWindowHost>
{
public:
    FancyZones(HINSTANCE hinstance, IFancyZonesSettings* settings) noexcept
        : m_hinstance(hinstance)
    {
        m_settings.attach(settings);
    }

    // IFancyZones
    IFACEMETHODIMP_(void) Run() noexcept;
    IFACEMETHODIMP_(void) Destroy() noexcept;

    // IFancyZonesCallback
    IFACEMETHODIMP_(bool) InMoveSize() noexcept { std::shared_lock readLock(m_lock); return m_inMoveSize; }
    IFACEMETHODIMP_(void) MoveSizeStart(HWND window, HMONITOR monitor, POINT const& ptScreen) noexcept;
    IFACEMETHODIMP_(void) MoveSizeUpdate(HMONITOR monitor, POINT const& ptScreen) noexcept;
    IFACEMETHODIMP_(void) MoveSizeEnd(HWND window, POINT const& ptScreen) noexcept;
    IFACEMETHODIMP_(void) VirtualDesktopChanged() noexcept;
    IFACEMETHODIMP_(bool) OnKeyDown(PKBDLLHOOKSTRUCT info) noexcept;

    // IZoneWindowHost
    IFACEMETHODIMP_(void) ToggleZoneViewers() noexcept;
    IFACEMETHODIMP_(void) MoveWindowsOnActiveZoneSetChange() noexcept;
    IFACEMETHODIMP_(COLORREF) GetZoneHighlightColor() noexcept
    {
        // Skip the leading # and convert to long
        const auto color = m_settings->GetSettings().zoneHightlightColor;
        const auto tmp = std::stol(color.substr(1), nullptr, 16);
        const auto nR = (tmp & 0xFF0000) >> 16;
        const auto nG = (tmp & 0xFF00) >> 8;
        const auto nB = (tmp & 0xFF);
        return RGB(nR, nG, nB);
    }

    LRESULT WndProc(HWND, UINT, WPARAM, LPARAM) noexcept;
    void OnDisplayChange(DisplayChangeType changeType) noexcept;
    void ShowZoneEditorForMonitor(HMONITOR monitor) noexcept;
    void AddZoneWindow(HMONITOR monitor, PCWSTR deviceId) noexcept;
    void MoveWindowIntoZoneByIndex(HWND window, int index) noexcept;

protected:
    static LRESULT CALLBACK s_WndProc(HWND, UINT, WPARAM, LPARAM) noexcept;

private:
    struct require_read_lock
    {
        template<typename T>
        require_read_lock(const std::shared_lock<T>& lock) { lock; }

        template<typename T>
        require_read_lock(const std::unique_lock<T>& lock) { lock; }
    };

    struct require_write_lock
    {
        template<typename T>
        require_write_lock(const std::unique_lock<T>& lock) { lock; }
    };

    void UpdateZoneWindows() noexcept;
    void MoveWindowsOnDisplayChange() noexcept;
    void UpdateDragState(require_write_lock) noexcept;
    void CycleActiveZoneSet(DWORD vkCode) noexcept;
    void OnSnapHotkey(DWORD vkCode) noexcept;
    void MoveSizeStartInternal(HWND window, HMONITOR monitor, POINT const& ptScreen, require_write_lock) noexcept;
    void MoveSizeEndInternal(HWND window, POINT const& ptScreen, require_write_lock) noexcept;
    void MoveSizeUpdateInternal(HMONITOR monitor, POINT const& ptScreen, require_write_lock) noexcept;

    const HINSTANCE m_hinstance{};
    static UINT WM_PRIV_VDCHANGED;

    HWND m_window{};
    HWND m_windowMoveSize{}; // The window that is being moved/sized
    bool m_editorsVisible{}; // Are we showing the zone editors?
    bool m_inMoveSize{};  // Whether or not a move/size operation is currently active
    bool m_dragEnabled{}; // True if we should be showing zone hints while dragging
    std::map<HMONITOR, winrt::com_ptr<IZoneWindow>> m_zoneWindowMap;
    winrt::com_ptr<IZoneWindow> m_zoneWindowMoveSize;
    winrt::com_ptr<IFancyZonesSettings> m_settings;
    GUID m_currentVirtualDesktopId{};
    mutable std::shared_mutex m_lock;
};

UINT FancyZones::WM_PRIV_VDCHANGED = RegisterWindowMessage(L"{128c2cb0-6bdf-493e-abbe-f8705e04aa95}");

// IFancyZones
IFACEMETHODIMP_(void) FancyZones::Run() noexcept
{
    std::unique_lock writeLock(m_lock);

    WNDCLASSEXW wcex{};
    wcex.cbSize = sizeof(WNDCLASSEX);
    wcex.lpfnWndProc = s_WndProc;
    wcex.hInstance = m_hinstance;
    wcex.lpszClassName = L"SuperFancyZones";
    RegisterClassExW(&wcex);

    BufferedPaintInit();

    m_window = CreateWindowExW(WS_EX_TOOLWINDOW, L"SuperFancyZones", L"", WS_POPUP, 0, 0, 0, 0, nullptr, nullptr, m_hinstance, this);
    if (!m_window) return;

    RegisterHotKey(m_window, 1, MOD_WIN, VK_OEM_3);
    VirtualDesktopChanged();
}

// IFancyZones
IFACEMETHODIMP_(void) FancyZones::Destroy() noexcept
{
    std::unique_lock writeLock(m_lock);

    BufferedPaintUnInit();
    if (m_window)
    {
        DestroyWindow(m_window);
        m_window = nullptr;
    }
}

// IFancyZonesCallback
IFACEMETHODIMP_(void) FancyZones::MoveSizeStart(HWND window, HMONITOR monitor, POINT const& ptScreen) noexcept
{
    std::unique_lock writeLock(m_lock);
    MoveSizeStartInternal(window, monitor, ptScreen, writeLock);
}

// IFancyZonesCallback
IFACEMETHODIMP_(void) FancyZones::MoveSizeUpdate(HMONITOR monitor, POINT const& ptScreen) noexcept
{
    std::unique_lock writeLock(m_lock);
    MoveSizeUpdateInternal(monitor, ptScreen, writeLock);
}

// IFancyZonesCallback
IFACEMETHODIMP_(void) FancyZones::MoveSizeEnd(HWND window, POINT const& ptScreen) noexcept
{
    std::unique_lock writeLock(m_lock);
    MoveSizeEndInternal(window, ptScreen, writeLock);
}

// IFancyZonesCallback
IFACEMETHODIMP_(void) FancyZones::VirtualDesktopChanged() noexcept
{
    // VirtualDesktopChanged is called from another thread but results in new windows being created.
    // Jump over to the UI thread to handle it.
    PostMessage(m_window, WM_PRIV_VDCHANGED, 0, 0);
}

// IFancyZonesCallback
IFACEMETHODIMP_(bool) FancyZones::OnKeyDown(PKBDLLHOOKSTRUCT info) noexcept
{
    // Return true to swallow the keyboard event
    bool const shift = GetAsyncKeyState(VK_SHIFT) & 0x8000;
    bool const win = GetAsyncKeyState(VK_LWIN) & 0x8000;
    if (win && !shift)
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
        else if ((info->vkCode == VK_RIGHT) || (info->vkCode == VK_LEFT))
        {
            OnSnapHotkey(info->vkCode);
            return true;
        }
    }
    else if (m_inMoveSize && (info->vkCode >= '0') && (info->vkCode <= '9'))
    {
        CycleActiveZoneSet(info->vkCode);
        return true;
    }
    return false;
}

// IZoneWindowHost
IFACEMETHODIMP_(void) FancyZones::ToggleZoneViewers() noexcept
{
    bool alreadyVisible{};

    {
        std::unique_lock writeLock(m_lock);
        alreadyVisible = m_editorsVisible;
        m_editorsVisible = !alreadyVisible;
    }

    if (!alreadyVisible)
    {
        auto callback = [](HMONITOR monitor, HDC, RECT *, LPARAM data) -> BOOL
        {
            auto strongThis = reinterpret_cast<FancyZones*>(data);
            strongThis->ShowZoneEditorForMonitor(monitor);
            return TRUE;
        };
        EnumDisplayMonitors(nullptr, nullptr, callback, reinterpret_cast<LPARAM>(this));
    }
    else
    {
        std::shared_lock readLock(m_lock);
        for (auto iter : m_zoneWindowMap)
        {
            iter.second->HideZoneWindow();
        }
    }
}

// IZoneWindowHost
IFACEMETHODIMP_(void) FancyZones::MoveWindowsOnActiveZoneSetChange() noexcept
{
    if (m_settings->GetSettings().zoneSetChange_moveWindows)
    {
        MoveWindowsOnDisplayChange();
    }
}

LRESULT FancyZones::WndProc(HWND window, UINT message, WPARAM wparam, LPARAM lparam) noexcept
{
    switch (message)
    {
    case WM_HOTKEY:
    {
        if (wparam == 1)
        {
            ToggleZoneViewers();
        }
    }
    break;

    case WM_SETTINGCHANGE:
    {
        if (wparam == SPI_SETWORKAREA)
        {
            OnDisplayChange(DisplayChangeType::WorkArea);
        }
    }
    break;

    case WM_DISPLAYCHANGE:
    {
        OnDisplayChange(DisplayChangeType::DisplayChange);
    }
    break;

    default:
    {
        if (message == WM_PRIV_VDCHANGED)
        {
            OnDisplayChange(DisplayChangeType::VirtualDesktop);
        }
        else
        {
            return DefWindowProc(window, message, wparam, lparam);
        }
    }
    break;
    }
    return 0;
}

void FancyZones::OnDisplayChange(DisplayChangeType changeType) noexcept
{
    if (changeType == DisplayChangeType::VirtualDesktop)
    {
        std::shared_lock readLock(m_lock);
        GUID currentVirtualDesktopId{};
        if (SUCCEEDED_LOG(RegistryHelpers::GetCurrentVirtualDesktop(&currentVirtualDesktopId)))
        {
            m_currentVirtualDesktopId = currentVirtualDesktopId;
        }
    }

    UpdateZoneWindows();

    if ((changeType == DisplayChangeType::WorkArea) || (changeType == DisplayChangeType::DisplayChange))
    {
        if (m_settings->GetSettings().displayChange_moveWindows)
        {
            MoveWindowsOnDisplayChange();
        }
    }
    else if (changeType == DisplayChangeType::VirtualDesktop)
    {
        if (m_settings->GetSettings().virtualDesktopChange_moveWindows)
        {
            MoveWindowsOnDisplayChange();
        }
    }
}

void FancyZones::ShowZoneEditorForMonitor(HMONITOR monitor) noexcept
{
    std::shared_lock readLock(m_lock);

    auto iter = m_zoneWindowMap.find(monitor);
    if (iter != m_zoneWindowMap.end())
    {
        bool const activate = MonitorFromPoint(POINT(), MONITOR_DEFAULTTOPRIMARY) == monitor;
        iter->second->ShowZoneWindow(activate, false /*fadeIn*/);
    }
}

void FancyZones::AddZoneWindow(HMONITOR monitor, PCWSTR deviceId) noexcept
{
    std::unique_lock writeLock(m_lock);
    wil::unique_cotaskmem_string virtualDesktopId;
    if (SUCCEEDED_LOG(StringFromCLSID(m_currentVirtualDesktopId, &virtualDesktopId)))
    {
        const bool flash = m_settings->GetSettings().virtualDesktopChange_flashZones;
        if (auto zoneWindow = MakeZoneWindow(this, m_hinstance, monitor, deviceId, virtualDesktopId.get(), flash))
        {
            m_zoneWindowMap[monitor] = std::move(zoneWindow);
        }
    }
}

void FancyZones::MoveWindowIntoZoneByIndex(HWND window, int index) noexcept
{
    std::shared_lock readLock(m_lock);
    if (window != m_windowMoveSize)
    {
        if (const HMONITOR monitor = MonitorFromWindow(window, MONITOR_DEFAULTTONULL))
        {
            auto iter = m_zoneWindowMap.find(monitor);
            if (iter != m_zoneWindowMap.end())
            {
                iter->second->MoveWindowIntoZoneByIndex(window, index);
            }
        }
    }
}

LRESULT CALLBACK FancyZones::s_WndProc(HWND window, UINT message, WPARAM wparam, LPARAM lparam) noexcept
{
    auto thisRef = reinterpret_cast<FancyZones*>(GetWindowLongPtr(window, GWLP_USERDATA));
    if (!thisRef && (message == WM_CREATE))
    {
        const auto createStruct = reinterpret_cast<LPCREATESTRUCT>(lparam);
        thisRef = reinterpret_cast<FancyZones*>(createStruct->lpCreateParams);
        SetWindowLongPtr(window, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(thisRef));
    }

    return thisRef ? thisRef->WndProc(window, message, wparam, lparam) :
        DefWindowProc(window, message, wparam, lparam);
}

void FancyZones::UpdateZoneWindows() noexcept
{
    auto callback = [](HMONITOR monitor, HDC, RECT *, LPARAM data) -> BOOL
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
                if (WI_IsFlagSet(displayDevice.StateFlags, DISPLAY_DEVICE_MIRRORING_DRIVER))
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

                auto strongThis = reinterpret_cast<FancyZones*>(data);
                strongThis->AddZoneWindow(monitor, deviceId);
            }
        }
        return TRUE;
    };

    {
        std::unique_lock writeLock(m_lock);
        m_zoneWindowMap.clear();
    }
    EnumDisplayMonitors(nullptr, nullptr, callback, reinterpret_cast<LPARAM>(this));
}

void FancyZones::MoveWindowsOnDisplayChange() noexcept
{
    auto callback = [](HWND window, LPARAM data) -> BOOL
    {
        int i = static_cast<int>(reinterpret_cast<UINT_PTR>(::GetProp(window, ZONE_STAMP)));
        if (i != 0)
        {
            // i is off by 1 since 0 is special.
            auto strongThis = reinterpret_cast<FancyZones*>(data);
            strongThis->MoveWindowIntoZoneByIndex(window, i-1);
        }
        return TRUE;
    };
    EnumWindows(callback, reinterpret_cast<LPARAM>(this));
}

void FancyZones::UpdateDragState(require_write_lock) noexcept
{
    const bool shift = GetAsyncKeyState(VK_SHIFT) & 0x8000;
    m_dragEnabled = m_settings->GetSettings().shiftDrag ? shift : !shift;
}

void FancyZones::CycleActiveZoneSet(DWORD vkCode) noexcept
{
    if (const HWND window = GetForegroundWindow())
    {
        if (const HMONITOR monitor = MonitorFromWindow(window, MONITOR_DEFAULTTONULL))
        {
            std::shared_lock readLock(m_lock);
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
    if (const HWND window = GetForegroundWindow())
    {
        if (const HMONITOR monitor = MonitorFromWindow(window, MONITOR_DEFAULTTONULL))
        {
            std::shared_lock readLock(m_lock);
            auto iter = m_zoneWindowMap.find(monitor);
            if (iter != m_zoneWindowMap.end())
            {
                iter->second->MoveWindowIntoZoneByDirection(window, vkCode);
            }
        }
    }
}

void FancyZones::MoveSizeStartInternal(HWND window, HMONITOR monitor, POINT const& ptScreen, require_write_lock writeLock) noexcept
{
    // Only enter move/size if the cursor is in the titlebar.
    // This prevents resize from triggering zones.
    RECT windowRect{};
    ::GetWindowRect(window, &windowRect);

    TITLEBARINFO titlebarInfo{ sizeof(titlebarInfo) };
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

            UpdateDragState(writeLock);

            if (m_dragEnabled)
            {
                m_zoneWindowMoveSize = iter->second;
                m_zoneWindowMoveSize->MoveSizeEnter(window, m_dragEnabled);
            }
            else if (m_zoneWindowMoveSize)
            {
                m_zoneWindowMoveSize->MoveSizeCancel();
                m_zoneWindowMoveSize = nullptr;
            }
        }
    }
}

void FancyZones::MoveSizeEndInternal(HWND window, POINT const& ptScreen, require_write_lock) noexcept
{
    m_inMoveSize = false;
    m_dragEnabled = false;
    m_windowMoveSize = nullptr;
    if (m_zoneWindowMoveSize)
    {
        auto zoneWindow = std::move(m_zoneWindowMoveSize);
        zoneWindow->MoveSizeEnd(window, ptScreen);
    }
    else
    {
        ::RemoveProp(window, ZONE_STAMP);
    }
}

void FancyZones::MoveSizeUpdateInternal(HMONITOR monitor, POINT const& ptScreen, require_write_lock writeLock) noexcept
{
    if (m_inMoveSize)
    {
        UpdateDragState(writeLock);

        if (m_zoneWindowMoveSize)
        {
            if (!m_dragEnabled)
            {
                auto zoneWindow = std::move(m_zoneWindowMoveSize);
                zoneWindow->MoveSizeCancel();
            }
            else
            {
                auto iter = m_zoneWindowMap.find(monitor);
                if (iter != m_zoneWindowMap.end())
                {
                    if (iter->second != m_zoneWindowMoveSize)
                    {
                        // The drag has moved to a different monitor.
                        auto const isDragEnabled = m_zoneWindowMoveSize->IsDragEnabled();
                        m_zoneWindowMoveSize->MoveSizeCancel();
                        m_zoneWindowMoveSize = iter->second;
                        m_zoneWindowMoveSize->MoveSizeEnter(m_windowMoveSize, isDragEnabled);
                    }
                    m_zoneWindowMoveSize->MoveSizeUpdate(ptScreen, m_dragEnabled);
                }
            }
        }
        else if (m_dragEnabled)
        {
            MoveSizeStartInternal(m_windowMoveSize, monitor, ptScreen, writeLock);
            MoveSizeUpdateInternal(monitor, ptScreen, writeLock);
        }
    }
}

winrt::com_ptr<IFancyZones> MakeFancyZones(HINSTANCE hinstance, IFancyZonesSettings* settings) noexcept
{
    return winrt::make_self<FancyZones>(hinstance, settings);
}
