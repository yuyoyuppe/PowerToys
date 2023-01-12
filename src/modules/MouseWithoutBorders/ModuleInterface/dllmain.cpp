#include "pch.h"
#include <interface/powertoy_module_interface.h>
#include <common/SettingsAPI/settings_objects.h>
#include <common/interop/shared_constants.h>
#include "trace.h"
#include <common/logger/logger.h>
#include <common/SettingsAPI/settings_helpers.h>

#include <common/utils/process_path.h>
#include <common/utils/resources.h>
#include <common/utils/winapi_error.h>

#include <filesystem>

BOOL APIENTRY DllMain(HMODULE /*hModule*/, DWORD ul_reason_for_call, LPVOID /*lpReserved*/)
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
        Trace::RegisterProvider();
        break;
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
        break;
    case DLL_PROCESS_DETACH:
        Trace::UnregisterProvider();
        break;
    }
    return TRUE;
}

const static wchar_t* MODULE_NAME = L"MouseWithoutBorders";
const static wchar_t* MODULE_DESC = L"A module to move your mouse across PCs.";

class MouseWithoutBorders : public PowertoyModuleIface
{
    std::wstring app_name;
    std::wstring app_key;

private:
    bool m_enabled = false;
    HANDLE send_telemetry_event;
    HANDLE m_hInvokeEvent;
    PROCESS_INFORMATION p_info;

    bool is_process_running()
    {
        return WaitForSingleObject(p_info.hProcess, 0) == WAIT_TIMEOUT;
    }

    void launch_process()
    {
        Logger::trace(L"Launching PowerToys MouseWithoutBorders process");
        const std::wstring application_path = L"modules\\MouseWithoutBorders\\PowerToys.MouseWithoutBorders.exe";
        STARTUPINFO info = { sizeof(info) };
        const unsigned long powertoys_pid = GetCurrentProcessId();
        std::wstring full_command_path = application_path + L" " + std::to_wstring(powertoys_pid);

        if (!CreateProcessW(application_path.c_str(), full_command_path.data(), NULL, NULL, true, NULL, NULL, NULL, &info, &p_info))
        {
            DWORD error = GetLastError();
            std::wstring message = L"PowerToys MouseWithoutBorders failed to start with error: ";
            message += std::to_wstring(error);
            Logger::error(message);
        }
    }

public:
    MouseWithoutBorders()
    {
        app_name = L"MouseWithoutBorders";
        app_key = app_name;
        std::filesystem::path logFilePath(PTSettingsHelper::get_module_save_folder_location(app_key));
        logFilePath.append(LogSettings::mouseWithoutBordersLogPath);
        Logger::init(LogSettings::mouseWithoutBordersLoggerName, logFilePath.wstring(), PTSettingsHelper::get_log_settings_file_location());
    };

    // Return the configured status for the gpo policy for the module
    virtual powertoys_gpo::gpo_rule_configured_t gpo_policy_enabled_configuration() override
    {
        return powertoys_gpo::getConfiguredMouseWithoutBordersEnabledValue();
    }

    virtual void destroy() override
    {
        TerminateProcess(p_info.hProcess, 1);
        delete this;
    }

    virtual const wchar_t* get_name() override
    {
        return MODULE_NAME;
    }

    virtual bool get_config(wchar_t* buffer, int* buffer_size) override
    {
        HINSTANCE hinstance = reinterpret_cast<HINSTANCE>(&__ImageBase);

        PowerToysSettings::Settings settings(hinstance, get_name());
        settings.set_description(MODULE_DESC);

        return settings.serialize_to_buffer(buffer, buffer_size);
    }

    virtual const wchar_t* get_key() override
    {
        return app_key.c_str();
    }

    virtual void set_config(const wchar_t* config) override
    {
        try
        {
            // Parse the input JSON string.
            PowerToysSettings::PowerToyValues values =
                PowerToysSettings::PowerToyValues::from_json_string(config, get_key());

            // If you don't need to do any custom processing of the settings, proceed
            // to persists the values.
            values.save_to_settings_file();
        }
        catch (std::exception&)
        {
            // Improper JSON.
        }
    }

    virtual void enable()
    {
        Trace::MouseWithoutBorders::Enable(true);
        ResetEvent(send_telemetry_event);
        ResetEvent(m_hInvokeEvent);
        launch_process();
        m_enabled = true;
    };

    virtual void disable()
    {
        if (m_enabled)
        {
            Trace::MouseWithoutBorders::Enable(false);
            Logger::trace(L"Disabling MouseWithoutBorders...");
            ResetEvent(send_telemetry_event);
            ResetEvent(m_hInvokeEvent);

            Logger::trace(L"Signaled exit event for PowerToys MouseWithoutBorders.");
            TerminateProcess(p_info.hProcess, 1);

            CloseHandle(p_info.hProcess);
        }

        m_enabled = false;
    }

    virtual bool is_enabled() override
    {
        return m_enabled;
    }
};

extern "C" __declspec(dllexport) PowertoyModuleIface* __cdecl powertoy_create()
{
    return new MouseWithoutBorders();
}