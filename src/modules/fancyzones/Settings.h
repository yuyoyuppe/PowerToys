#pragma once

#define ZONE_STAMP L"FancyZones_zone"

struct Settings
{
    // The values specified here are the defaults.
    bool shiftDrag = false;
    bool displayChange_moveWindows = false;
    bool virtualDesktopChange_moveWindows = false;
    bool virtualDesktopChange_flashZones = true;
    bool zoneSetChange_moveWindows = false;
    bool overrideSnapHotkeys = true;
    std::wstring zoneHightlightColor = L"#0078D7";
};

interface __declspec(uuid("{BA4E77C4-6F44-4C5D-93D3-CBDE880495C2}")) IFancyZonesSettings : public IUnknown
{
    IFACEMETHOD_(bool, GetConfig)(_Out_ PCWSTR* config) = 0;
    IFACEMETHOD_(void, SetConfig)(PCWSTR config) = 0;
    IFACEMETHOD_(Settings, GetSettings)() = 0;
};

winrt::com_ptr<IFancyZonesSettings> MakeFancyZonesSettings(HINSTANCE hinstance, PCWSTR config) noexcept;