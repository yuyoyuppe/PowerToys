#pragma once
#include "FancyZones.h"

interface __declspec(uuid("{7F017528-8110-4FB3-BE41-F472969C2560}")) IZoneWindow : public IUnknown
{
    IFACEMETHOD(ShowZoneWindow)(bool activate) = 0;
    IFACEMETHOD(HideZoneWindow)() = 0;
    IFACEMETHOD(MoveSizeEnter)(_In_ HWND window, POINT ptScreen, DragMode dragMode) = 0;
    IFACEMETHOD(MoveSizeExit)(_In_ HWND window, POINT ptScreen) = 0;
    IFACEMETHOD(MoveSizeUpdate)(POINT ptScreen, DragMode dragMode) = 0;
    IFACEMETHOD(MoveSizeCancel)() = 0;
    IFACEMETHOD_(DragMode, GetDragMode)() = 0;
    IFACEMETHOD_(void, MoveWindowIntoZoneByIndex)(_In_ HWND window, int index) = 0;
    IFACEMETHOD_(void, MoveWindowIntoZoneByDirection)(_In_ HWND window, DWORD vkCode) = 0;
    IFACEMETHOD_(void, OnDisplayChange)(DisplayChangeType type) = 0;
    IFACEMETHOD_(void, CycleActiveZoneSet)(DWORD vkCode) = 0;
};

Microsoft::WRL::ComPtr<IZoneWindow> MakeZoneWindow(_In_ HMONITOR monitor, _In_ PCWSTR deviceId) noexcept;