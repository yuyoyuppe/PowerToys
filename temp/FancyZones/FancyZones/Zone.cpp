#include "stdafx.h"

using namespace Microsoft::WRL;

class Zone WrlFinal : public RuntimeClass<
    RuntimeClassFlags<Microsoft::WRL::ClassicCom>,
    IZone>
{
public:
    Zone(RECT zoneRect) :
        m_zoneRect(zoneRect)
    {
    }

    IFACEMETHODIMP_(RECT) GetZoneRect() noexcept { return m_zoneRect; }
    IFACEMETHODIMP_(bool) ContainsWindow(_In_ HWND window) noexcept;
    IFACEMETHODIMP_(void) AddWindowToZone(_In_ HWND window, _In_ HWND zoneWindow, bool stampZone) noexcept;
    IFACEMETHODIMP_(void) RemoveWindowFromZone(_In_ HWND window, bool restoreSize) noexcept;
    IFACEMETHODIMP_(void) SetId(size_t id) noexcept { m_id = id; }
    IFACEMETHODIMP_(size_t) GetId() noexcept { return m_id; }

private:
    void SizeWindowToZone(_In_ HWND window, _In_ HWND zoneWindow) noexcept;
    void StampZone(_In_ HWND window, bool stamp) noexcept;

    RECT m_zoneRect{};
    size_t m_id{};
    std::map<HWND, RECT> m_windows{};
};

IFACEMETHODIMP_(bool) Zone::ContainsWindow(_In_ HWND window) noexcept
{
    return (m_windows.find(window) != m_windows.end());
}

IFACEMETHODIMP_(void) Zone::AddWindowToZone(_In_ HWND window, _In_ HWND zoneWindow, bool stampZone) noexcept
{
    WINDOWPLACEMENT placement;
    ::GetWindowPlacement(window, &placement);
    ::GetWindowRect(window, &placement.rcNormalPosition);
    m_windows.emplace(std::pair<HWND, RECT>(window, placement.rcNormalPosition));

    SizeWindowToZone(window, zoneWindow);
    if (stampZone)
    {
        StampZone(window, true);
    }
}

IFACEMETHODIMP_(void) Zone::RemoveWindowFromZone(_In_ HWND window, bool restoreSize) noexcept
{
    auto iter = m_windows.find(window);
    if (iter != m_windows.end())
    {
        m_windows.erase(iter);
        StampZone(window, false);
    }
}

void Zone::SizeWindowToZone(_In_ HWND window, _In_ HWND zoneWindow) noexcept
{
    // Take care of 1px border
    RECT zoneRect = m_zoneRect;

    RECT windowRect{};
    ::GetWindowRect(window, &windowRect);

    RECT frameRect{};
    if (SUCCEEDED(DwmGetWindowAttribute(window, DWMWA_EXTENDED_FRAME_BOUNDS, &frameRect, sizeof(frameRect))))
    {
        zoneRect.bottom -= (frameRect.bottom - windowRect.bottom);
        zoneRect.right -= (frameRect.right - windowRect.right);
        zoneRect.left -= (frameRect.left - windowRect.left);
    }

    // Map to screen coords
    MapWindowRect(zoneWindow, nullptr, &zoneRect);
    ::SetWindowPos(window, nullptr, zoneRect.left, zoneRect.top, zoneRect.right - zoneRect.left, zoneRect.bottom - zoneRect.top, SWP_NOZORDER | SWP_ASYNCWINDOWPOS | SWP_NOACTIVATE | SWP_NOSENDCHANGING);
}

void Zone::StampZone(_In_ HWND window, bool stamp) noexcept
{
    if (stamp)
    {
        SetProp(window, ZONE_STAMP, reinterpret_cast<HANDLE>(m_id));
    }
    else
    {
        RemoveProp(window, ZONE_STAMP);
    }
}

ComPtr<IZone> MakeZone(RECT zoneRect) noexcept
{
    return Make<Zone>(zoneRect);
}