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
#include <Lmcons.h>
#include <sddl.h>
#include <common/utils/processApi.h>

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

bool GetUserSid(const wchar_t* username, PSID& sid)
{
    DWORD sidSize = 0;
    DWORD domainNameSize = 0;
    SID_NAME_USE sidNameUse;

    LookupAccountName(nullptr, username, nullptr, &sidSize, nullptr, &domainNameSize, &sidNameUse);
    if (GetLastError() != ERROR_INSUFFICIENT_BUFFER)
    {
        Logger::error("Failed to get buffer sizes");
        return false;
    }

    sid = LocalAlloc(LPTR, sidSize);
    LPWSTR domainName = static_cast<LPWSTR>(LocalAlloc(LPTR, domainNameSize * sizeof(wchar_t)));

    if (!LookupAccountNameW(nullptr, username, sid, &sidSize, domainName, &domainNameSize, &sidNameUse))
    {
        Logger::error("Failed to lookup account name");
        LocalFree(sid);
        LocalFree(domainName);
        return false;
    }

    LocalFree(domainName);
    return true;
}

std::wstring GetCurrentUserSid()
{
    wchar_t username[UNLEN + 1];
    DWORD usernameSize = UNLEN + 1;

    std::wstring result;
    if (!GetUserNameW(username, &usernameSize))
    {
        Logger::error("Failed to get the current user name");
        return result;
    }

    PSID sid;
    if (GetUserSid(username, sid))
    {
        LPWSTR sidString;
        if (ConvertSidToStringSid(sid, &sidString))
        {
            result = sidString;
            LocalFree(sidString);
        }
        LocalFree(sid);
    }
    else
    {
        Logger::error(L"Failed to get SID for user \"");
    }

    return result;
}

std::wstring escapeDoubleQuotes(const std::wstring& input)
{
    std::wstring output;
    output.reserve(input.size());

    for (const wchar_t& ch : input)
    {
        if (ch == L'"')
        {
            output += L'\\';
        }
        output += ch;
    }

    return output;
}

const static wchar_t* MODULE_NAME = L"MouseWithoutBorders";
const static wchar_t* MODULE_DESC = L"A module to move your mouse across PCs.";
const static wchar_t* SERVICE_NAME = L"Mouse Without Borders service";

class MouseWithoutBorders : public PowertoyModuleIface
{
    std::wstring app_name;
    std::wstring app_key;

private:
    bool m_enabled = false;
    HANDLE send_telemetry_event;
    HANDLE m_hInvokeEvent;
    PROCESS_INFORMATION p_info;

    bool is_enabled_by_default() const override
    {
        return false;
    }

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

        if (!CreateProcessW(application_path.c_str(), full_command_path.data(), nullptr, nullptr, true, {}, nullptr, nullptr, &info, &p_info))
        {
            DWORD error = GetLastError();
            std::wstring message = L"PowerToys MouseWithoutBorders failed to start with error: ";
            message += std::to_wstring(error);
            Logger::error(message);
        }
    }

    void toggle_service_registration(const bool enable)
    {
        if (enable)
        {
            SC_HANDLE schSCManager = OpenSCManagerW(nullptr, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS);
            if (schSCManager == nullptr)
            {
                Logger::error(L"Couldn't open scm manager");
                return;
            }

            const auto servicePath = std::filesystem::current_path() / "modules/MouseWithoutBorders/PowerToys.MouseWithoutBordersService.exe";

            // Pass localappdata of the current user to the service
            PWSTR cLocalAppPath;
            winrt::check_hresult(SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, NULL, &cLocalAppPath));
            CoTaskMemFree(cLocalAppPath);

            std::wstring localAppPath{ cLocalAppPath };
            std::wstring binaryWithArgsPath = L"\"";
            binaryWithArgsPath += servicePath;
            binaryWithArgsPath += L"\" ";
            binaryWithArgsPath += escapeDoubleQuotes(localAppPath);

            SC_HANDLE schService = CreateServiceW(
                schSCManager,
                SERVICE_NAME,
                SERVICE_NAME,
                SERVICE_ALL_ACCESS,
                SERVICE_WIN32_OWN_PROCESS,
                SERVICE_AUTO_START,
                SERVICE_ERROR_NORMAL,
                binaryWithArgsPath.c_str(),
                nullptr,
                nullptr,
                nullptr,
                nullptr,
                nullptr);

            if (schService == nullptr)
            {
                Logger::error(L"Failed to create service");
                CloseServiceHandle(schSCManager);
                return;
            }

            // Set up the security descriptor to allow non-elevated users to start the service
            PSECURITY_DESCRIPTOR pSD = nullptr;
            ULONG szSD = 0;
            std::wstring securityDescriptor = L"D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)(A;;CR;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)(A;;RPWPDTLO;;;";

            securityDescriptor += GetCurrentUserSid();
            securityDescriptor += L")S:(AU;FA;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;WD)";

            if (!ConvertStringSecurityDescriptorToSecurityDescriptorW(
                    securityDescriptor.c_str(),
                    SDDL_REVISION_1,
                    &pSD,
                    &szSD))
            {
                Logger::error(L"Failed to convert security descriptor string");
                CloseServiceHandle(schService);
                CloseServiceHandle(schSCManager);
                return;
            }

            if (!SetServiceObjectSecurity(schService, DACL_SECURITY_INFORMATION, pSD))
            {
                Logger::error("Failed to set service object security, error: ");
            }

            LocalFree(pSD);
            CloseServiceHandle(schService);
            CloseServiceHandle(schSCManager);
        }
        else
        {
            std::wstring service_stop_cmd = L"sc stop ";
            service_stop_cmd += '"';
            service_stop_cmd += SERVICE_NAME;
            service_stop_cmd += '"';
            _wsystem(service_stop_cmd.c_str());

            std::wstring service_delete_cmd = L"sc delete ";
            service_delete_cmd += '"';
            service_delete_cmd += SERVICE_NAME;
            service_delete_cmd += '"';
            _wsystem(service_delete_cmd.c_str());
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

    void shutdown_processes()
    {
        const auto services = getProcessHandlesByName(L"PowerToys.MouseWithoutBordersService.exe", PROCESS_TERMINATE);
        for (const auto& svc : services)
            TerminateProcess(svc.get(), 0);

        const auto apps = getProcessHandlesByName(L"PowerToys.MouseWithoutBorders.exe", PROCESS_TERMINATE);
        for (const auto& app : apps)
            TerminateProcess(app.get(), 0);
    }

    virtual void destroy() override
    {
        shutdown_processes();
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
        toggle_service_registration(true);
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

            shutdown_processes();

            CloseHandle(p_info.hProcess);
            toggle_service_registration(false);
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