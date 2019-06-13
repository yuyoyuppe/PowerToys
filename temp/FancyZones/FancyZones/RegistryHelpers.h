#pragma once

namespace RegistryHelpers
{
    static PCWSTR REG_SETTINGS = L"Software\\SuperFancyZones";

    inline PCWSTR GetKey(_In_opt_ PCWSTR monitorId, PWSTR key, size_t keyLength)
    {
        if (monitorId)
        {
            StringCchPrintf(key, keyLength, L"%s\\%s", REG_SETTINGS, monitorId);
        }
        else
        {
            StringCchPrintf(key, keyLength, L"%s", REG_SETTINGS);
        }
        return key;
    }

    inline HKEY OpenKey(_In_opt_ PCWSTR monitorId)
    {
        HKEY hkey;
        wchar_t key[256];
        GetKey(monitorId, key, ARRAYSIZE(key));
        if (RegOpenKeyExW(HKEY_CURRENT_USER, key, 0, KEY_ALL_ACCESS, &hkey) == ERROR_SUCCESS)
        {
            return hkey;
        }
        return nullptr;
    }

    inline HKEY CreateKey(_In_ PCWSTR monitorId)
    {
        HKEY hkey;
        wchar_t key[256]{};
        GetKey(monitorId, key, ARRAYSIZE(key));
        if (RegCreateKeyExW(HKEY_CURRENT_USER, key, 0, nullptr, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, nullptr, &hkey, nullptr) == ERROR_SUCCESS)
        {
            return hkey;
        }
        return nullptr;
    }

    template<typename t>
    inline void GetValue(_In_ PCWSTR monitorId, PCWSTR setting, t* value, DWORD size)
    {
        // XXXX: revisit Get/Set, can do better than create/open
        // Stop using USValue
        wchar_t key[256]{};
        GetKey(monitorId, key, ARRAYSIZE(key));
        SHRegGetUSValueW(key, setting, nullptr, value, &size, FALSE, nullptr, 0);
    }

    inline void GetPath(PCWSTR monitorId, PCWSTR setting, PWSTR value)
    {
        wchar_t key[256]{};
        GetKey(monitorId, key, ARRAYSIZE(key));
        SHRegGetPathW(HKEY_CURRENT_USER, key, setting, value, 0);
    }

    inline void SetPath(PCWSTR monitorId, PCWSTR setting, PWSTR value)
    {
        wchar_t key[256]{};
        GetKey(monitorId, key, ARRAYSIZE(key));
        SHRegSetPathW(HKEY_CURRENT_USER, key, setting, value, 0);
    }

    template<typename t>
    inline void SetValue(_In_ PCWSTR monitorId, PCWSTR setting, t value, DWORD size)
    {
        // XXXX: revisit Get/Set, can do better than create/open
        // Stop using USValue
        wchar_t key[256]{};
        GetKey(monitorId, key, ARRAYSIZE(key));
        SHRegSetUSValueW(key, setting, REG_BINARY, &value, size, SHREGSET_FORCE_HKCU);
    }

    inline void DeleteZoneSet(_In_ PCWSTR monitorId, GUID guid)
    {
        PWSTR zoneSetId;
        if (SUCCEEDED(StringFromCLSID(guid, &zoneSetId)))
        {
            wchar_t key[256]{};
            GetKey(monitorId, key, ARRAYSIZE(key));
            SHDeleteValueW(HKEY_CURRENT_USER, key, zoneSetId);
            CoTaskMemFree(zoneSetId);
        }
    }

    inline void DeleteAllZoneSets(_In_ PCWSTR monitorId)
    {
        wchar_t key[256]{};
        GetKey(monitorId, key, ARRAYSIZE(key));
        SHDeleteKey(HKEY_CURRENT_USER, key);
    }

    inline HRESULT GetCurrentVirtualDesktop(GUID* id)
    {
        DWORD sessionId;
        ProcessIdToSessionId(GetCurrentProcessId(), &sessionId);

        wchar_t sessionKeyPath[256]{};
        HRESULT hr = StringCchPrintfW(
            sessionKeyPath,
            ARRAYSIZE(sessionKeyPath),
            L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\SessionInfo\\%d\\VirtualDesktops",
            sessionId);
        if (SUCCEEDED(hr))
        {
            HKEY key{};
            GUID value{};
            if (RegOpenKeyExW(HKEY_CURRENT_USER, sessionKeyPath, 0, KEY_ALL_ACCESS, &key) == ERROR_SUCCESS)
            {
                DWORD size = sizeof(value);
                if (RegQueryValueExW(key, L"CurrentVirtualDesktop", 0, nullptr, reinterpret_cast<BYTE*>(&value), &size) == ERROR_SUCCESS)
                {
                    *id = value;
                    RegCloseKey(key);
                    return S_OK;
                }
                RegCloseKey(key);
            }
        }
        return hr;
    }
}