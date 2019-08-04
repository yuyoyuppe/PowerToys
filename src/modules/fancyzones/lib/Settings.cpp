#include "pch.h"
#include <common/settings_objects.h>

struct FancyZonesSettings : winrt::implements<FancyZonesSettings, IFancyZonesSettings>
{
public:
    FancyZonesSettings(HINSTANCE hinstance, PCWSTR name)
        : m_hinstance(hinstance)
        , m_name(name)
    {
        LoadSettings(name, true /*fromFile*/);
    }

    IFACEMETHODIMP_(bool) GetConfig(_Out_ PCWSTR* config) noexcept;
    IFACEMETHODIMP_(void) SetConfig(PCWSTR config) noexcept;
    IFACEMETHODIMP_(Settings) GetSettings() noexcept { return m_settings; }

private:
    void LoadSettings(PCWSTR config, bool fromFile) noexcept;
    void SaveSettings() noexcept;

    const HINSTANCE m_hinstance;
    PCWSTR m_name{};

    Settings m_settings;

    struct
    {
        PCWSTR name;
        bool* value;
        int resourceId;
    } m_configBools[6] = {
        { L"fancyzones_shiftDrag", &m_settings.shiftDrag, IDS_SETTING_DESCRIPTION_SHIFTDRAG },
        { L"fancyzones_overrideSnapHotkeys", &m_settings.overrideSnapHotkeys, IDS_SETTING_DESCRIPTION_OVERRIDE_SNAP_HOTKEYS },
        { L"fancyzones_displayChange_moveWindows", &m_settings.displayChange_moveWindows, IDS_SETTING_DESCRIPTION_DISPLAYCHANGE_MOVEWINDOWS },
        { L"fancyzones_zoneSetChange_moveWindows", &m_settings.zoneSetChange_moveWindows, IDS_SETTING_DESCRIPTION_ZONESETCHANGE_MOVEWINDOWS },
        { L"fancyzones_virtualDesktopChange_moveWindows", &m_settings.virtualDesktopChange_moveWindows, IDS_SETTING_DESCRIPTION_VIRTUALDESKTOPCHANGE_MOVEWINDOWS },
        { L"fancyzones_virtualDesktopChange_flashZones", &m_settings.virtualDesktopChange_flashZones, IDS_SETTING_DESCRIPTION_VIRTUALDESKTOPCHANGE_FLASHZONES },
    };

    struct
    {
        PCWSTR name;
        std::wstring* value;
        int resourceId;
    } m_configStrings[1] = {
        { L"fancyzones_zoneHighlightColor", &m_settings.zoneHightlightColor, IDS_SETTING_DESCRIPTION_ZONEHIGHLIGHTCOLOR },
    };

};

IFACEMETHODIMP_(bool) FancyZonesSettings::GetConfig(_Out_ PCWSTR* config) noexcept
{
    PowerToysSettings::Settings settings(m_name, L"Helps organize your windows.");

    settings.set_icon_key(L"pt-fancy-zones");

    for (auto const& setting : m_configBools)
    {
        wchar_t description[256];
        LoadString(m_hinstance, setting.resourceId, description, ARRAYSIZE(description));
        settings.add_property(PowerToysSettings::BoolTogglePropertySetting(setting.name, description, *setting.value));
    }

    for (auto const& setting : m_configStrings)
    {
        wchar_t description[256];
        LoadString(m_hinstance, setting.resourceId, description, ARRAYSIZE(description));
        settings.add_property(PowerToysSettings::ColorPickerPropertySetting(setting.name, description, *setting.value));
    }

    *config = settings.to_allocated_cstring();

    return true;
}

IFACEMETHODIMP_(void) FancyZonesSettings::SetConfig(PCWSTR config) noexcept try
{
    LoadSettings(config, false /*fromFile*/);
    SaveSettings();
}
CATCH_LOG();

void FancyZonesSettings::LoadSettings(PCWSTR config, bool fromFile) noexcept try
{
    PowerToysSettings::PowerToyValues values = fromFile ?
        PowerToysSettings::PowerToyValues::load_from_settings_file(m_name) :
        PowerToysSettings::PowerToyValues::from_json_string(config);

    for (auto const& setting : m_configBools)
    {
        if (values.is_bool_value(setting.name))
        {
            *setting.value = values.get_bool_value(setting.name);
        }
    }

    for (auto const& setting : m_configStrings)
    {
        if (values.is_string_value(setting.name))
        {
            *setting.value = values.get_string_value(setting.name);
        }
    }
}
CATCH_LOG();

void FancyZonesSettings::SaveSettings() noexcept try
{
    PowerToysSettings::PowerToyValues values(m_name);

    for (auto const& setting : m_configBools)
    {
        values.add_property_value(PowerToysSettings::BoolSettingsValue(setting.name, *setting.value));
    }

    for (auto const& setting : m_configStrings)
    {
        values.add_property_value(PowerToysSettings::StringSettingsValue(setting.name, *setting.value));
    }

    values.save_to_settings_file();
}
CATCH_LOG();

winrt::com_ptr<IFancyZonesSettings> MakeFancyZonesSettings(HINSTANCE hinstance, PCWSTR name) noexcept
{
    return winrt::make_self<FancyZonesSettings>(hinstance, name);
}