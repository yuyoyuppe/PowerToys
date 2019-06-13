#pragma once

#include "Zone.h"

enum class ZoneSetLayout
{
    Grid,
    Row,
    Focus,
    Custom
};

interface __declspec(uuid("{E4839EB7-669D-49CF-84A9-71A2DFD851A3}")) IZoneSet : public IUnknown
{
    IFACEMETHOD_(GUID, GetId)() = 0;
    IFACEMETHOD(AddZone)(_In_ Microsoft::WRL::ComPtr<IZone> zone, bool front) = 0;
    IFACEMETHOD(RemoveZone)(_In_ Microsoft::WRL::ComPtr<IZone> zone) = 0;
    IFACEMETHOD_(Microsoft::WRL::ComPtr<IZone>, ZoneFromPoint)(POINT pt) = 0;
    IFACEMETHOD_(Microsoft::WRL::ComPtr<IZone>, ZoneFromWindow)(_In_ HWND window) = 0;
    IFACEMETHOD_(std::vector<Microsoft::WRL::ComPtr<IZone>>, GetZones)() = 0;
    IFACEMETHOD_(ZoneSetLayout, GetLayout)() = 0;
    IFACEMETHOD_(int, GetInnerPadding)() = 0;
    IFACEMETHOD_(Microsoft::WRL::ComPtr<IZoneSet>, MakeCustomClone)() = 0;
    IFACEMETHOD_(void, Save)() = 0;
    IFACEMETHOD_(void, MoveZoneToFront)(_In_ Microsoft::WRL::ComPtr<IZone> zone) = 0;
    IFACEMETHOD_(void, MoveZoneToBack)(_In_ Microsoft::WRL::ComPtr<IZone> zone) = 0;
    IFACEMETHOD_(void, MoveWindowIntoZoneByIndex)(_In_ HWND window, _In_ HWND zoneWindow, int index) = 0;
    IFACEMETHOD_(void, MoveWindowIntoZoneByDirection)(_In_ HWND window, _In_ HWND zoneWindow, DWORD vkCode) = 0;
    IFACEMETHOD_(void, MoveSizeExit)(_In_ HWND window, _In_ HWND zoneWindow, _In_ POINT ptClient) = 0;
};

#define VERSION_PERSISTEDDATA 0x0000F00D
struct ZoneSetPersistedData
{
    DWORD version{VERSION_PERSISTEDDATA};
    DWORD zoneCount{};
    ZoneSetLayout layout{};
    DWORD paddingInner{};
    DWORD paddingOuter{};
    RECT zones[25]{};
};

struct ZoneSetConfig
{
    ZoneSetConfig(
        GUID idIn,
        _In_ HMONITOR monitorIn,
        PCWSTR workAreaIn,
        ZoneSetLayout layoutIn,
        int zoneCountIn,
        int paddingOuterIn,
        int paddingInnerIn) noexcept :
        id(idIn),
        monitor(monitorIn),
        workArea(workAreaIn),
        layout(layoutIn),
        zoneCount(zoneCountIn),
        paddingOuter(paddingOuterIn),
        paddingInner(paddingInnerIn)
    {
    }

    GUID id{};
    HMONITOR monitor{};
    PCWSTR workArea{};
    ZoneSetLayout layout{};
    int zoneCount{};
    int paddingOuter{};
    int paddingInner{};
    bool custom{};
};

Microsoft::WRL::ComPtr<IZoneSet> MakeZoneSet(ZoneSetConfig const& config) noexcept;