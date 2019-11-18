#include "pch.h"
#include "common/dpi_aware.h"
#include "common/on_thread_executor.h"

#include "common/wmi_query.h"
#include "WmiHelpers.h"

#include <cinttypes>
#include <unordered_set>

struct FancyZones : public winrt::implements<FancyZones, IFancyZones, IFancyZonesCallback, IZoneWindowHost>
{
public:
    FancyZones(HINSTANCE hinstance, IFancyZonesSettings* settings) noexcept
        : m_hinstance(hinstance)
        , m_settings(settings)
        , m_wmiConnection(wmi_connection::initialize())
    {
        m_settings->SetCallback(this);
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
    IFACEMETHODIMP_(void) WindowCreated(HWND window) noexcept;
    IFACEMETHODIMP_(bool) OnKeyDown(PKBDLLHOOKSTRUCT info) noexcept;
    IFACEMETHODIMP_(void) ToggleEditor() noexcept;
    IFACEMETHODIMP_(void) SettingsChanged() noexcept;

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

    mutable std::shared_mutex m_lock;
    HWND m_window{};
    HWND m_windowMoveSize{}; // The window that is being moved/sized
    bool m_editorsVisible{}; // Are we showing the zone editors?
    bool m_inMoveSize{};  // Whether or not a move/size operation is currently active
    bool m_dragEnabled{}; // True if we should be showing zone hints while dragging
    std::map<HMONITOR, winrt::com_ptr<IZoneWindow>> m_zoneWindowMap; // Map of monitor to ZoneWindow (one per monitor)
    winrt::com_ptr<IZoneWindow> m_zoneWindowMoveSize; // "Active" ZoneWindow, where the move/size is happening. Will update as drag moves between monitors.
    IFancyZonesSettings* m_settings{};
    GUID m_currentVirtualDesktopId{};
    wil::unique_handle m_terminateEditorEvent;

    OnThreadExecutor m_dpiUnawareThread;

    wmi_connection m_wmiConnection;

    static UINT WM_PRIV_VDCHANGED;
    static UINT WM_PRIV_EDITOR;

    enum class EditorExitKind : byte
    {
        Exit,
        Terminate
    };
};

UINT FancyZones::WM_PRIV_VDCHANGED = RegisterWindowMessage(L"{128c2cb0-6bdf-493e-abbe-f8705e04aa95}");
UINT FancyZones::WM_PRIV_EDITOR = RegisterWindowMessage(L"{87543824-7080-4e91-9d9c-0404642fc7b6}");

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

    RegisterHotKey(m_window, 1, m_settings->GetSettings().editorHotkey.get_modifiers(), m_settings->GetSettings().editorHotkey.get_code());
    VirtualDesktopChanged();

    m_dpiUnawareThread.submit(OnThreadExecutor::task_t{[]{
        SetThreadDpiAwarenessContext(DPI_AWARENESS_CONTEXT_UNAWARE);
        SetThreadDpiHostingBehavior(DPI_HOSTING_BEHAVIOR_MIXED);
    }}).wait();
}

// IFancyZones
IFACEMETHODIMP_(void) FancyZones::Destroy() noexcept
{
    std::unique_lock writeLock(m_lock);
    m_zoneWindowMap.clear();
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
IFACEMETHODIMP_(void) FancyZones::WindowCreated(HWND window) noexcept
{
    if (m_settings->GetSettings().appLastZone_moveWindows)
    {
        auto processPath = get_process_path(window);
        if (!processPath.empty()) 
        {
            INT zoneIndex = -1;
            LRESULT res = RegistryHelpers::GetAppLastZone(window, processPath.data(), &zoneIndex);
            if ((res == ERROR_SUCCESS) && (zoneIndex != -1))
            {
                MoveWindowIntoZoneByIndex(window, zoneIndex);
            }
        }
    }
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
                Trace::FancyZones::OnKeyDown(info->vkCode, win, ctrl, false /* inMoveSize */);
                CycleActiveZoneSet(info->vkCode);
                return true;
            }
        }
        else if ((info->vkCode == VK_RIGHT) || (info->vkCode == VK_LEFT))
        {
            Trace::FancyZones::OnKeyDown(info->vkCode, win, ctrl, false /* inMoveSize */);
            OnSnapHotkey(info->vkCode);
            return true;
        }
    }
    else if (m_inMoveSize && (info->vkCode >= '0') && (info->vkCode <= '9'))
    {
        Trace::FancyZones::OnKeyDown(info->vkCode, win, false /* control */, true/* inMoveSize */);
        CycleActiveZoneSet(info->vkCode);
        return true;
    }
    return false;
}

// IFancyZonesCallback
void FancyZones::ToggleEditor() noexcept
{
    {
        std::shared_lock readLock(m_lock);
        if (m_terminateEditorEvent)
        {
            SetEvent(m_terminateEditorEvent.get());
            return;
        }
    }

    {
        std::unique_lock writeLock(m_lock);
        m_terminateEditorEvent.reset(CreateEvent(nullptr, true, false, nullptr));
    }

    HMONITOR monitor{};
    HWND foregroundWindow{};

    UINT dpi_x = DPIAware::DEFAULT_DPI;
    UINT dpi_y = DPIAware::DEFAULT_DPI;

    const bool use_cursorpos_editor_startupscreen = m_settings->GetSettings().use_cursorpos_editor_startupscreen;
    POINT currentCursorPos{};
    if (use_cursorpos_editor_startupscreen)
    {
        GetCursorPos(&currentCursorPos);
        monitor = MonitorFromPoint(currentCursorPos, MONITOR_DEFAULTTOPRIMARY);
    }
    else
    {
        foregroundWindow = GetForegroundWindow();
        monitor = MonitorFromWindow(foregroundWindow, MONITOR_DEFAULTTOPRIMARY);
    }


    if (!monitor)
    {
        return;
    }

    std::shared_lock readLock(m_lock);
    auto iter = m_zoneWindowMap.find(monitor);
    if (iter == m_zoneWindowMap.end())
    {
        return;
    }

    MONITORINFOEX mi;
    mi.cbSize = sizeof(mi);

    m_dpiUnawareThread.submit(OnThreadExecutor::task_t{[&]{
        GetMonitorInfo(monitor, &mi);
    }}).wait();

    if(use_cursorpos_editor_startupscreen)
    {
        DPIAware::GetScreenDPIForPoint(currentCursorPos, dpi_x, dpi_y);
    }
    else
    {
        DPIAware::GetScreenDPIForWindow(foregroundWindow, dpi_x, dpi_y);
    }

    const auto taskbar_x_offset = MulDiv(mi.rcWork.left - mi.rcMonitor.left, DPIAware::DEFAULT_DPI, dpi_x);
    const auto taskbar_y_offset = MulDiv(mi.rcWork.top - mi.rcMonitor.top, DPIAware::DEFAULT_DPI, dpi_y);
    
    // Do not scale window params by the dpi, that will be done in the editor - see LayoutModel.Apply
    const auto x = mi.rcMonitor.left + taskbar_x_offset;
    const auto y = mi.rcMonitor.top + taskbar_y_offset;
    const auto width = mi.rcWork.right - mi.rcWork.left;
    const auto height = mi.rcWork.bottom - mi.rcWork.top;
    const std::wstring editorLocation = 
        std::to_wstring(x) + L"_" +
        std::to_wstring(y) + L"_" +
        std::to_wstring(width) + L"_" +
        std::to_wstring(height);

    const std::wstring params =
        iter->second->UniqueId() + L" " +
        std::to_wstring(iter->second->ActiveZoneSet()->LayoutId()) + L" " +
        std::to_wstring(reinterpret_cast<UINT_PTR>(monitor)) + L" " +
        editorLocation + L" " +
        iter->second->WorkAreaKey() + L" " +
        std::to_wstring(static_cast<float>(dpi_x) / DPIAware::DEFAULT_DPI);

    SHELLEXECUTEINFO sei{ sizeof(sei) };
    sei.fMask = { SEE_MASK_NOCLOSEPROCESS | SEE_MASK_FLAG_NO_UI };
    sei.lpFile = L"modules\\FancyZonesEditor.exe";
    sei.lpParameters = params.c_str();
    sei.nShow = SW_SHOWNORMAL;
    ShellExecuteEx(&sei);

    // Launch the editor on a background thread
    // Wait for the editor's process to exit
    // Post back to the main thread to update
    std::thread waitForEditorThread([window = m_window, processHandle = sei.hProcess, terminateEditorEvent = m_terminateEditorEvent.get()]()
    {
        HANDLE waitEvents[2] = { processHandle, terminateEditorEvent };
        auto result = WaitForMultipleObjects(2, waitEvents, false, INFINITE);
        if (result == WAIT_OBJECT_0 + 0)
        {
            // Editor exited
            // Update any changes it may have made
            PostMessage(window, WM_PRIV_EDITOR, 0, static_cast<LPARAM>(EditorExitKind::Exit));
        }
        else if (result == WAIT_OBJECT_0 + 1)
        {
            // User hit Win+~ while editor is already running
            // Shut it down
            TerminateProcess(processHandle, 2);
            PostMessage(window, WM_PRIV_EDITOR, 0, static_cast<LPARAM>(EditorExitKind::Terminate));
        }
        CloseHandle(processHandle);
    });

    waitForEditorThread.detach();
}

void FancyZones::SettingsChanged() noexcept
{
    // Update the hotkey
    UnregisterHotKey(m_window, 1);
    RegisterHotKey(m_window, 1, m_settings->GetSettings().editorHotkey.get_modifiers(), m_settings->GetSettings().editorHotkey.get_code());
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
    Trace::FancyZones::ToggleZoneViewers(!alreadyVisible);

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
            if (m_settings->GetSettings().use_standalone_editor)
            {
                ToggleEditor();
            }
            else
            {
                ToggleZoneViewers();
            }
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
        else if (message == WM_PRIV_EDITOR)
        {
            if (lparam == static_cast<LPARAM>(EditorExitKind::Exit))
            {
                // Don't reload settings if we terminated the editor
                OnDisplayChange(DisplayChangeType::Editor);
            }

            {
                // Clean up the event either way
                std::unique_lock writeLock(m_lock);
                m_terminateEditorEvent.release();
            }
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
        // Explorer persists this value to the registry on a per session basis but only after
        // the first virtual desktop switch happens. If the user hasn't switched virtual desktops in this session
        // then this value will be empty. This means loading the first virtual desktop's configuration can be
        // funky the first time we load up at boot since the user will not have switched virtual desktops yet.
        std::shared_lock readLock(m_lock);
        GUID currentVirtualDesktopId{};
        if (SUCCEEDED(RegistryHelpers::GetCurrentVirtualDesktop(&currentVirtualDesktopId)))
        {
            m_currentVirtualDesktopId = currentVirtualDesktopId;
        }
        else
        {
            // TODO: Use the previous "Desktop 1" fallback
            // Need to maintain a map of desktop name to virtual desktop uuid
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
    else if (changeType == DisplayChangeType::Editor)
    {
        if (m_settings->GetSettings().zoneSetChange_moveWindows)
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
        const bool flash = m_settings->GetSettings().zoneSetChange_flashZones;
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
    {
      std::unique_lock writeLock(m_lock);
      m_zoneWindowMap.clear();
    }

    struct device
    {
      std::wstring _deviceID;
      HMONITOR _handle;

      std::optional<std::wstring_view> get_generated_part_of_id() const
      {
          if (const auto sharpPos = _deviceID.find(L'#'); sharpPos != std::wstring::npos)
          {
            return std::wstring_view{_deviceID.data() + sharpPos + 1, size(_deviceID) - sharpPos - 1};
          }
          return std::nullopt;
      }
    };
    auto callback = [](HMONITOR monitor, HDC, RECT *, LPARAM captured_ptr) -> BOOL
    {
        const auto doContinueEnumeration = TRUE;
        MONITORINFOEX mi;
        mi.cbSize = sizeof(mi);
        if (GetMonitorInfo(monitor, &mi) == FALSE)
        {
          return doContinueEnumeration;
        }
        DISPLAY_DEVICE displayDevice = { sizeof(displayDevice) };

        if (EnumDisplayDevices(mi.szDevice, 0, &displayDevice, 1))
        {
            if (WI_IsFlagSet(displayDevice.StateFlags, DISPLAY_DEVICE_MIRRORING_DRIVER))
            {
                return doContinueEnumeration;
            }
        }
        auto devices = reinterpret_cast<std::vector<device> *>(captured_ptr);
        device newDevice{{}, monitor};
        if (displayDevice.DeviceID[0] != L'\0')
        {
            std::array<wchar_t, 256> parsedId{};
            ParseDeviceId(displayDevice.DeviceID, parsedId.data(), size(parsedId));
            newDevice._deviceID = parsedId.data();
        }
        devices->emplace_back(std::move(newDevice));
        return doContinueEnumeration;
    };
    std::vector<device> devices;
    EnumDisplayMonitors(nullptr, nullptr, callback, reinterpret_cast<LPARAM>(&devices));
    std::vector<WmiMonitorID> monitor_infos;
    try
    {
        m_wmiConnection.select_all(L"WmiMonitorID", [&](std::wstring_view xml_obj) {
          monitor_infos.emplace_back(parse_monitorID_from_dtd(xml_obj));
        });
    }
    catch(...) {}

    if(GetSystemMetrics(SM_REMOTESESSION))
    {
        // When running in a remote session, devices have "Default_Monitor#<transient_part>" DeviceIDS.
        // Since a remote session can have multiple devices, we can't just assign them a single "REMOTE" ID.
        // We'll assign them unique numeric IDs instead.
        for (size_t i = 0; i < size(devices); ++i)
        {
            wchar_t remoteDeviceID[32];
            swprintf_s(remoteDeviceID, L"REMOTE%zu", i);
            AddZoneWindow(devices[i]._handle, remoteDeviceID);
        }
        return;
    }
    // If EnumDisplayMonitors returned a duplicate data, it's an API bug and we should workaround it
    const bool enumDisplayMonitorsDublicateData = [&]{
        std::unordered_set<std::wstring_view> discoveredDeviceIDs;
        for (const auto & dev : devices)
        {
            const auto generatedPartOfID = dev.get_generated_part_of_id();
            if (!generatedPartOfID.has_value())
            {
                continue;
            }
            const auto [_, hasUniqueID] = discoveredDeviceIDs.emplace(*generatedPartOfID);
            if (!hasUniqueID)
            {
                return true;
            }
        }
        return false;
    }();


    for (size_t i = 0; i < size(devices); ++i)
    {
        const auto & dev = devices[i];
        const auto generatedPartOfID = dev.get_generated_part_of_id();
        // In case DeviceID doesn't contain valid data or we can't rely on it, since EnumDisplayMonitors
        // returned a duplicate results, we try using WMI data or fallback to UNKNOWN + numeric ID.
        if (!generatedPartOfID.has_value() || enumDisplayMonitorsDublicateData)
        {
            const bool hasWmiData = i < size(monitor_infos);
            if (hasWmiData)
            {
                AddZoneWindow(dev._handle, monitor_infos[i].hardware_id().c_str());
            }
            else
            {
                // no WMI data - likely we're running inside a VM
                wchar_t localDeviceID[32];
                swprintf_s(localDeviceID, L"UNKNOWN%zu", i);
                AddZoneWindow(dev._handle, localDeviceID);
            }
            continue;
        }
        // The happy path is to associate generated parts of DeviceIDs with a WMIMonitorID.InstanceName
        // And use WMI IDs instead. We can't get WMI data if we're running inside a VM, but at this point we've
        // avoided API bugs so can somewhat rely on DeviceID.
        bool associatedWithWMI = false;
        for (const auto & wmi_info : monitor_infos)
        {
            associatedWithWMI = wmi_info._instance_name.find(*generatedPartOfID) != std::wstring::npos;
            if (associatedWithWMI)
            {
                AddZoneWindow(dev._handle, wmi_info.hardware_id().c_str());
                break;
            }
        }
        if (!associatedWithWMI)
        {
            AddZoneWindow(dev._handle, generatedPartOfID->data());
        }
    }
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
    if (const HWND window = get_filtered_active_window())
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
    if (const HWND window = get_filtered_active_window())
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
    // Only enter move/size if the cursor is inside the window rect by a certain padding.
    // This prevents resize from triggering zones.
    RECT windowRect{};
    ::GetWindowRect(window, &windowRect);

    const auto padding_x = 8;
    const auto padding_y = 6;
    windowRect.top += padding_y;
    windowRect.left += padding_x;
    windowRect.right -= padding_x;
    windowRect.bottom -= padding_y;

    if (PtInRect(&windowRect, ptScreen) == FALSE)
    {
        return;
    }

    m_inMoveSize = true;

    auto iter = m_zoneWindowMap.find(monitor);
    if (iter == end(m_zoneWindowMap))
    {
        return;
    }

    m_windowMoveSize = window;

    // This updates m_dragEnabled depending on if the shift key is being held down.
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

        auto processPath = get_process_path(window);
        if (!processPath.empty())
        {
            RegistryHelpers::SaveAppLastZone(window, processPath.data(), -1);
        }
    }
}

void FancyZones::MoveSizeUpdateInternal(HMONITOR monitor, POINT const& ptScreen, require_write_lock writeLock) noexcept
{
    if (m_inMoveSize)
    {
        // This updates m_dragEnabled depending on if the shift key is being held down.
        UpdateDragState(writeLock);

        if (m_zoneWindowMoveSize)
        {
            // Update the ZoneWindow already handling move/size
            if (!m_dragEnabled)
            {
                // Drag got disabled, tell it to cancel and clear out m_zoneWindowMoveSize
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
            // We'll get here if the user presses/releases shift while dragging.
            // Restart the drag on the ZoneWindow that m_windowMoveSize is on
            MoveSizeStartInternal(m_windowMoveSize, monitor, ptScreen, writeLock);
            MoveSizeUpdateInternal(monitor, ptScreen, writeLock);
        }
    }
}

winrt::com_ptr<IFancyZones> MakeFancyZones(HINSTANCE hinstance, IFancyZonesSettings* settings) noexcept
{
    return winrt::make_self<FancyZones>(hinstance, settings);
}