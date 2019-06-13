#pragma once

#define ZONE_STAMP L"FancyZones_zone"

enum class DragMode
{
    None,
    Normal,
    SetWindowPos,
};

#define SETTINGS_VERSION_0 0x0000F00D
#define SETTINGS_VERSION_1 0x0001F00D
#define SETTINGS_VERSION_CURRENT SETTINGS_VERSION_1
struct Settings
{
    // Version 0
    DWORD version{SETTINGS_VERSION_CURRENT};
    DragMode dragMode{DragMode::Normal};
    bool displayChange_moveWindows = false;
    bool virtualDesktopChange_moveWindows = false;
    bool overrideSnapHotkeys = true;

    // Version 1
    bool virtualDesktopChange_wallpaper = false; // Deprecated
    bool virtualDesktopChange_flashZones = true;
    bool colorfulZones = false;
    bool zoneSetChange_moveWindows = false;
    bool zoneSetChange_flashZones = true;
};

interface __declspec(uuid("{BA4E77C4-6F44-4C5D-93D3-CBDE880495C2}")) IFancyZonesSettings : public IUnknown
{
    IFACEMETHOD_(Settings, GetSettings)() = 0;
};

Microsoft::WRL::ComPtr<IFancyZonesSettings> MakeFancyZonesSettings() noexcept;