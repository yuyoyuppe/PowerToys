#pragma once

interface __declspec(uuid("{8228E934-B6EF-402A-9892-15A1441BF8B0}")) IZone : public IUnknown
{
    IFACEMETHOD_(RECT, GetZoneRect)() = 0;
    IFACEMETHOD_(bool, ContainsWindow)(_In_ HWND window) = 0;
    IFACEMETHOD_(void, AddWindowToZone)(_In_ HWND window, _In_ HWND zoneWindow, bool stampZone) = 0;
    IFACEMETHOD_(void, RemoveWindowFromZone)(_In_ HWND window, bool restoreSize) = 0;
    IFACEMETHOD_(void, SetId)(size_t id) = 0;
    IFACEMETHOD_(size_t, GetId)() = 0;
};

Microsoft::WRL::ComPtr<IZone> MakeZone(RECT zoneRect) noexcept;
