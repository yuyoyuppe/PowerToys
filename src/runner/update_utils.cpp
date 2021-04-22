#include "pch.h"

#include "Generated Files/resource.h"

#include "action_runner_utils.h"
#include "update_state.h"
#include "update_utils.h"

#include <common/updating/installer.h>
#include <common/updating/http_client.h>
#include <common/updating/updating.h>
#include <common/utils/resources.h>
#include <common/utils/timeutil.h>
#include <runner/general_settings.h>

auto Strings = create_notifications_strings();

namespace
{
    constexpr int64_t UPDATE_CHECK_INTERVAL_MINUTES = 60 * 24;
    constexpr int64_t UPDATE_CHECK_AFTER_FAILED_INTERVAL_MINUTES = 60 * 2;

    const size_t MAX_DOWNLOAD_ATTEMPTS = 3;
}

bool start_msi_uninstallation_sequence()
{
    const auto package_path = updating::get_msi_package_path();

    if (package_path.empty())
    {
        // No MSI version detected
        return true;
    }

    if (!updating::offer_msi_uninstallation(Strings))
    {
        // User declined to uninstall or opted for "Don't show again"
        return false;
    }
    auto sei = launch_action_runner(L"-uninstall_msi");

    WaitForSingleObject(sei.hProcess, INFINITE);
    DWORD exit_code = 0;
    GetExitCodeProcess(sei.hProcess, &exit_code);
    CloseHandle(sei.hProcess);
    return exit_code == 0;
}

using namespace updating;

bool could_be_costly_connection()
{
    using namespace winrt::Windows::Networking::Connectivity;
    ConnectionProfile internetConnectionProfile = NetworkInformation::GetInternetConnectionProfile();
    return internetConnectionProfile.IsWwanConnectionProfile();
}

std::optional<std::filesystem::path> create_download_path()
{
    auto installer_download_path = get_pending_updates_path();
    std::error_code ec;
    std::filesystem::create_directories(installer_download_path, ec);
    return !ec ? std::optional{ installer_download_path } : std::nullopt;
}

std::future<bool> download_new_version(const new_version_download_info& new_version, const notifications::strings& strings)
{
    auto installer_download_path = create_download_path();
    if (!installer_download_path)
    {
        co_return false;
    }
    
    *installer_download_path /= new_version.installer_filename;

    bool download_success = false;
    for (size_t i = 0; i < MAX_DOWNLOAD_ATTEMPTS; ++i)
    {
        try
        {
            http::HttpClient client;
            co_await client.download(new_version.installer_download_url, *installer_download_path);
            download_success = true;
            break;
        }
        catch (...)
        {
            // reattempt to download or do nothing
        }
    }
    co_return download_success;
}

void proceed_with_update(const new_version_download_info& download_info, const bool download_update)
{
    if (!download_update)
    {
        notifications::show_visit_github(download_info, Strings);
        return;
    }

    if (!notifications::show_confirm_update(download_info, Strings))
    {
        return;
    }

    if (download_new_version(download_info, Strings).get())
    {
        std::wstring args{ UPDATE_NOW_LAUNCH_STAGE1_CMDARG };
        args += L' ';
        args += download_info.installer_filename;
        launch_action_runner(args.c_str());
    }
    else
    {
        notifications::show_install_error(download_info, Strings);
    }
}

void github_update_worker()
{
    for (;;)
    {
        auto state = UpdateState::read();
        int64_t sleep_minutes_till_next_update = 0;
        if (state.github_update_last_checked_date.has_value())
        {
            int64_t last_checked_minutes_ago = timeutil::diff::in_minutes(timeutil::now(), *state.github_update_last_checked_date);
            if (last_checked_minutes_ago < 0)
            {
                last_checked_minutes_ago = UPDATE_CHECK_INTERVAL_MINUTES;
            }
            sleep_minutes_till_next_update = max(0, UPDATE_CHECK_INTERVAL_MINUTES - last_checked_minutes_ago);
        }

        std::this_thread::sleep_for(std::chrono::minutes{ sleep_minutes_till_next_update });

        const bool download_update = !could_be_costly_connection() && get_general_settings().downloadUpdatesAutomatically;
        bool version_info_obtained = false;
        try
        {
            const auto version_info = get_github_version_info_async(Strings).get();
            version_info_obtained = version_info.has_value();
            if (version_info_obtained && std::holds_alternative<new_version_download_info>(*version_info))
            {
                const auto download_info = std::get<new_version_download_info>(*version_info);
                proceed_with_update(download_info, download_update);
            }
        }
        catch (...)
        {
            // Couldn't autoupdate
        }

        if (version_info_obtained)
        {
            UpdateState::store([](UpdateState& state) {
                state.github_update_last_checked_date.emplace(timeutil::now());
            });
        }
        else
        {
            std::this_thread::sleep_for(std::chrono::minutes{ UPDATE_CHECK_AFTER_FAILED_INTERVAL_MINUTES });
        }
    }
}

std::optional<updating::github_version_info> check_for_updates()
{
    try
    {
        auto version_info = get_github_version_info_async(Strings).get();
        if (!version_info)
        {
            notifications::show_unavailable(Strings, std::move(version_info.error()));
            return std::nullopt;
        }

        if (std::holds_alternative<updating::version_up_to_date>(*version_info))
        {
            notifications::show_unavailable(Strings, Strings.GITHUB_NEW_VERSION_UP_TO_DATE);
            return std::move(*version_info);
        }

        auto new_version = std::get<updating::new_version_download_info>(*version_info);
        return std::move(new_version);
    }
    catch (...)
    {
        // Couldn't autoupdate
    }
    return std::nullopt;
}

bool launch_pending_update()
{
    try
    {
        auto update_state = UpdateState::read();
        if (update_state.pending_update)
        {
            UpdateState::store([](UpdateState& state) {
                state.pending_update = false;
                state.pending_installer_filename = {};
            });
            std::wstring args{ UPDATE_NOW_LAUNCH_STAGE1_START_PT_CMDARG };
            args += L' ';
            args += update_state.pending_installer_filename;

            launch_action_runner(args.c_str());
            return true;
        }
    }
    catch (...)
    {
    }
    return false;
}
