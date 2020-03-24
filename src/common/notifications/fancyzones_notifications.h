#pragma once

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#include <Windows.h>

namespace
{
    const inline wchar_t CANT_DRAG_ELEVATED_DONT_SHOW_AGAIN_REGISTRY_PATH[] = LR"(SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\DontShowMeThisDialogAgain\{e16ea82f-6d94-4f30-bb02-d6d911588afd})";
}

inline bool disable_cant_drag_elevated_warning()
{
    HKEY key{};
    if (RegCreateKeyExW(HKEY_CURRENT_USER,
                        CANT_DRAG_ELEVATED_DONT_SHOW_AGAIN_REGISTRY_PATH,
                        0,
                        nullptr,
                        REG_OPTION_NON_VOLATILE,
                        KEY_ALL_ACCESS,
                        nullptr,
                        &key,
                        nullptr) != ERROR_SUCCESS)
    {
        return false;
    }
    DWORD value = 1;
    if (RegSetValueExW(key, nullptr, 0, REG_DWORD, reinterpret_cast<const BYTE*>(&value), sizeof(value)) != ERROR_SUCCESS)
    {
        RegCloseKey(key);
        return false;
    }
    RegCloseKey(key);
    return true;
}

inline bool is_cant_drag_elevated_warning_disabled()
{
    HKEY key{};
    if (RegOpenKeyExW(HKEY_CURRENT_USER,
                      CANT_DRAG_ELEVATED_DONT_SHOW_AGAIN_REGISTRY_PATH,
                      0,
                      KEY_READ,
                      &key) != ERROR_SUCCESS)
    {
        return false;
    }
    DWORD value = 0;
    DWORD value_size = sizeof(value);
    if (RegGetValueW(
            key,
            nullptr,
            nullptr,
            RRF_RT_DWORD,
            nullptr,
            &value,
            &value_size) != ERROR_SUCCESS)
    {
        RegCloseKey(key);
        return false;
    }
    RegCloseKey(key);
    return value == 1;
}