#include "stdafx.h"

using namespace Microsoft::WRL;

UINT const UM_TRAY_NOTIFICATION = (WM_USER + 0x100);

class FancyZonesSettings WrlFinal : public RuntimeClass<
    RuntimeClassFlags<Microsoft::WRL::ClassicCom>,
    IFancyZonesSettings>
{
public:
    FancyZonesSettings()
    {
        InitTrayIcon();
        LoadSettings();
    }

    IFACEMETHODIMP_(Settings) GetSettings() noexcept { return m_settings; }

protected:
    static LRESULT CALLBACK s_WndProc(HWND window, UINT message, WPARAM wparam, LPARAM lparam) noexcept;

private:
    void LoadSettings() noexcept;
    void SaveSettings() noexcept;
    void InitTrayIcon() noexcept;
    void DoExit() noexcept;
    void AddShellNotifyIcon() noexcept;
    LRESULT WndProc(UINT message, WPARAM wparam, LPARAM lparam) noexcept;

    HWND m_window{};
    Settings m_settings{};
    bool m_settingsLoaded{};
    UINT s_taskbarCreated{};
};

void FancyZonesSettings::LoadSettings() noexcept
{
    HKEY key = RegistryHelpers::OpenKey(nullptr);
    if (key)
    {
        Settings settings{};
        DWORD size = sizeof(settings);
        if (RegQueryValueExW(key, L"Settings", 0, nullptr, reinterpret_cast<BYTE*>(&settings), &size) == ERROR_SUCCESS)
        {
            if (settings.version == SETTINGS_VERSION_CURRENT)
            {
                m_settings = settings;
                m_settingsLoaded = true;
            }
            else if (settings.version == SETTINGS_VERSION_0)
            {
                // Just added a few bools, no extra work necessary
                m_settings = settings;
                m_settings.version = SETTINGS_VERSION_CURRENT;
                m_settingsLoaded = true;
            }
        }
        RegCloseKey(key);
    }
}

void FancyZonesSettings::SaveSettings() noexcept
{
    HKEY key = RegistryHelpers::CreateKey(nullptr);
    if (key)
    {
        RegSetValueExW(key, L"Settings", 0, REG_BINARY, reinterpret_cast<BYTE*>(&m_settings), sizeof(m_settings));
        RegCloseKey(key);
    }
}

void FancyZonesSettings::InitTrayIcon() noexcept
{
    WNDCLASSEXW wcex{};
    wcex.cbSize = sizeof(WNDCLASSEX);
    wcex.lpfnWndProc = s_WndProc;
    wcex.hInstance = g_app->GetHInstance();
    wcex.lpszClassName = L"SuperFancyZones_TrayWnd";
    RegisterClassExW(&wcex);

    m_window = CreateWindowExW(WS_EX_TOOLWINDOW, L"SuperFancyZones_TrayWnd", L"", WS_POPUP, 0, 0, 0, 0, nullptr, nullptr, g_app->GetHInstance(), this);
}

void FancyZonesSettings::DoExit() noexcept
{
    PostQuitMessage(0);

    NOTIFYICONDATA nid = { sizeof(nid) };
    nid.uID = (UINT)(void*)this;
    nid.hWnd = m_window;
    Shell_NotifyIcon(NIM_DELETE, &nid);
    m_window = nullptr;
}

void FancyZonesSettings::AddShellNotifyIcon() noexcept
{
    NOTIFYICONDATA nid = { sizeof(nid) };
    nid.uID = (UINT)(void*)this;
    nid.hWnd = m_window;
    nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    nid.hIcon = (HICON)LoadImage(g_app->GetHInstance(), MAKEINTRESOURCE(IDI_FANCYZONES), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR);
    nid.uCallbackMessage = UM_TRAY_NOTIFICATION;
    StringCchCopy(nid.szTip, ARRAYSIZE(nid.szTip), L"SuperFancyZones");
    Shell_NotifyIcon(NIM_ADD, &nid);
}

LRESULT FancyZonesSettings::WndProc(UINT message, WPARAM wparam, LPARAM lparam) noexcept
{
    switch (message)
    {
        case WM_NCDESTROY:
        {
            ::DefWindowProc(m_window, message, wparam, lparam);
            SetWindowLongPtr(m_window, GWLP_USERDATA, 0);
        }
        break;

        case WM_CREATE:
        {
            s_taskbarCreated = RegisterWindowMessage(L"TaskbarCreated");
            AddShellNotifyIcon();
        }
        break;

        case UM_TRAY_NOTIFICATION:
            switch (lparam)
            {
                case WM_RBUTTONDOWN:
                case WM_LBUTTONDOWN:
                {
                    HMENU menu = LoadMenu(g_app->GetHInstance(), MAKEINTRESOURCE(IDC_FANCYZONES));
                    HMENU popup = GetSubMenu(menu, 0);

                    UINT const checked = MF_BYCOMMAND | MF_CHECKED;
                    switch (m_settings.dragMode)
                    {
                        case DragMode::None: CheckMenuItem(popup, IDM_DRAGMODE_NONE, checked); break;
                        case DragMode::Normal: CheckMenuItem(popup, IDM_DRAGMODE_NORMAL, checked); break;
                    }

                    CheckMenuItem(popup,
                        IDM_DISPLAYCHANGE_MOVEWINDOWS,
                        MF_BYCOMMAND | m_settings.displayChange_moveWindows ? MF_CHECKED : MF_UNCHECKED);

                    CheckMenuItem(popup,
                        IDM_VIRTUALDESKTOPCHANGE_MOVEWINDOWS,
                        MF_BYCOMMAND | m_settings.virtualDesktopChange_moveWindows ? MF_CHECKED : MF_UNCHECKED);

                    CheckMenuItem(popup,
                        IDM_VIRTUALDESKTOPCHANGE_FLASHZONES,
                        MF_BYCOMMAND | m_settings.virtualDesktopChange_flashZones ? MF_CHECKED : MF_UNCHECKED);

                    CheckMenuItem(popup,
                        IDM_ZONESETCHANGE_MOVEWINDOWS,
                        MF_BYCOMMAND | m_settings.zoneSetChange_moveWindows ? MF_CHECKED : MF_UNCHECKED);

                    CheckMenuItem(popup,
                        IDM_ZONESETCHANGE_FLASHZONES,
                        MF_BYCOMMAND | m_settings.zoneSetChange_flashZones ? MF_CHECKED : MF_UNCHECKED);

                    CheckMenuItem(popup,
                        IDM_OVERRIDE_SNAP_HOTKEYS,
                        MF_BYCOMMAND | m_settings.overrideSnapHotkeys ? MF_CHECKED : MF_UNCHECKED);

                    CheckMenuItem(popup,
                        IDM_COLORFUL_ZONES,
                        MF_BYCOMMAND | m_settings.colorfulZones ? MF_CHECKED : MF_UNCHECKED);

                    POINT pt;
                    GetCursorPos(&pt);

                    SetForegroundWindow(m_window);
                    TrackPopupMenu(
                        popup,
                        GetSystemMetrics(SM_MENUDROPALIGNMENT) ? TPM_RIGHTALIGN : TPM_LEFTALIGN |
                        TPM_BOTTOMALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON,
                        pt.x, pt.y,
                        0, m_window, nullptr);

                    DestroyMenu(popup);
                    DestroyMenu(menu);
                    break;
                }
            }
            break;

        case WM_COMMAND:
            switch (LOWORD(wparam))
            {
                case IDM_EXIT: DoExit(); break;
                case IDM_TOGGLEZONEVIEWER: g_app->ToggleZoneViewers(); break;

                case IDM_DRAGMODE_NONE: m_settings.dragMode = DragMode::None; SaveSettings();  break;
                case IDM_DRAGMODE_NORMAL: m_settings.dragMode = DragMode::Normal; SaveSettings(); break;

                case IDM_DISPLAYCHANGE_MOVEWINDOWS: m_settings.displayChange_moveWindows = !m_settings.displayChange_moveWindows; SaveSettings(); break;

                case IDM_VIRTUALDESKTOPCHANGE_MOVEWINDOWS: m_settings.virtualDesktopChange_moveWindows = !m_settings.virtualDesktopChange_moveWindows; SaveSettings(); break;
                case IDM_VIRTUALDESKTOPCHANGE_FLASHZONES: m_settings.virtualDesktopChange_flashZones = !m_settings.virtualDesktopChange_flashZones; SaveSettings(); break;

                case IDM_ZONESETCHANGE_MOVEWINDOWS: m_settings.zoneSetChange_moveWindows = !m_settings.zoneSetChange_moveWindows; SaveSettings(); break;
                case IDM_ZONESETCHANGE_FLASHZONES: m_settings.zoneSetChange_flashZones = !m_settings.zoneSetChange_flashZones; SaveSettings(); break;

                case IDM_OVERRIDE_SNAP_HOTKEYS: m_settings.overrideSnapHotkeys = !m_settings.overrideSnapHotkeys; SaveSettings(); break;
                case IDM_COLORFUL_ZONES: m_settings.colorfulZones = !m_settings.colorfulZones; SaveSettings(); break;
            }
            break;

        case WM_QUIT:
        {
            DestroyWindow(m_window);
        }
        break;

        case WM_DESTROY:
        {
            DoExit();
        }
        break;

        default:
        {
            if (message == s_taskbarCreated)
            {
                AddShellNotifyIcon();
            }
            else
            {
                return DefWindowProc(m_window, message, wparam, lparam);
            }
        }
    }
    return 0;
}

LRESULT CALLBACK FancyZonesSettings::s_WndProc(_In_ HWND window, UINT message, _In_ WPARAM wparam, _In_ LPARAM lparam) noexcept
{
    auto thisRef = reinterpret_cast<FancyZonesSettings*>(GetWindowLongPtr(window, GWLP_USERDATA));
    if ((thisRef == nullptr) && (message == WM_CREATE))
    {
        auto createStruct = reinterpret_cast<LPCREATESTRUCT>(lparam);
        thisRef = reinterpret_cast<FancyZonesSettings*>(createStruct->lpCreateParams);
        thisRef->m_window = window;
        SetWindowLongPtr(window, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(thisRef));
    }

    return (thisRef != nullptr) ? thisRef->WndProc(message, wparam, lparam) :
        DefWindowProc(window, message, wparam, lparam);
}

ComPtr<IFancyZonesSettings> MakeFancyZonesSettings() noexcept
{
    return Make<FancyZonesSettings>();
}