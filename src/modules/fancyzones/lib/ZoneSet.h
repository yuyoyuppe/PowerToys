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
    IFACEMETHOD(AddZone)(winrt::com_ptr<IZone> zone, bool front) = 0;
    IFACEMETHOD(RemoveZone)(winrt::com_ptr<IZone> zone) = 0;
    IFACEMETHOD_(winrt::com_ptr<IZone>, ZoneFromPoint)(POINT pt) = 0;
    IFACEMETHOD_(winrt::com_ptr<IZone>, ZoneFromWindow)(HWND window) = 0;
    IFACEMETHOD_(int, GetZoneIndexFromWindow)(HWND window) = 0;
    IFACEMETHOD_(std::vector<winrt::com_ptr<IZone>>, GetZones)() = 0;
    IFACEMETHOD_(ZoneSetLayout, GetLayout)() = 0;
    IFACEMETHOD_(int, GetInnerPadding)() = 0;
    IFACEMETHOD_(winrt::com_ptr<IZoneSet>, MakeCustomClone)() = 0;
    IFACEMETHOD_(void, Save)() = 0;
    IFACEMETHOD_(void, MoveZoneToFront)(winrt::com_ptr<IZone> zone) = 0;
    IFACEMETHOD_(void, MoveZoneToBack)(winrt::com_ptr<IZone> zone) = 0;
    IFACEMETHOD_(void, MoveWindowIntoZoneByIndex)(HWND window, HWND zoneWindow, int index) = 0;
    IFACEMETHOD_(void, MoveWindowIntoZoneByDirection)(HWND window, HWND zoneWindow, DWORD vkCode) = 0;
    IFACEMETHOD_(void, MoveSizeEnd)(HWND window, HWND zoneWindow, POINT ptClient) = 0;
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
        HMONITOR monitorIn,
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

winrt::com_ptr<IZoneSet> MakeZoneSet(ZoneSetConfig const& config) noexcept;