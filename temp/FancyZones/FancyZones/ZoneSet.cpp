#include "stdafx.h"

using namespace Microsoft::WRL;

class ZoneSet WrlFinal : public RuntimeClass<
    RuntimeClassFlags<Microsoft::WRL::ClassicCom>,
    IZoneSet>
{
public:
    ZoneSet(ZoneSetConfig const& config) : m_config(config)
    {
        if (config.zoneCount > 0)
        {
            InitialPopulateZones();
        }
    }

    ZoneSet(ZoneSetConfig const& config, _In_ std::vector<ComPtr<IZone>> zones) :
        m_config(config),
        m_zones(zones)
    {
    }

    IFACEMETHODIMP_(GUID) GetId() noexcept { return m_config.id; }
    IFACEMETHODIMP AddZone(_In_ ComPtr<IZone> zone, bool front) noexcept;
    IFACEMETHODIMP RemoveZone(_In_ ComPtr<IZone> zone) noexcept;
    IFACEMETHODIMP_(ComPtr<IZone>) ZoneFromPoint(POINT pt) noexcept;
    IFACEMETHODIMP_(ComPtr<IZone>) ZoneFromWindow(_In_ HWND window) noexcept;
    IFACEMETHODIMP_(std::vector<ComPtr<IZone>>) GetZones() noexcept { return m_zones; }
    IFACEMETHODIMP_(ZoneSetLayout) GetLayout() noexcept { return m_config.layout; }
    IFACEMETHODIMP_(int) GetInnerPadding() noexcept { return m_config.paddingInner; }
    IFACEMETHODIMP_(ComPtr<IZoneSet>) MakeCustomClone() noexcept;
    IFACEMETHODIMP_(void) Save() noexcept;
    IFACEMETHODIMP_(void) MoveZoneToFront(_In_ ComPtr<IZone> zone) noexcept;
    IFACEMETHODIMP_(void) MoveZoneToBack(_In_ ComPtr<IZone> zone) noexcept;
    IFACEMETHODIMP_(void) MoveWindowIntoZoneByIndex(_In_ HWND window, _In_ HWND zoneWindow, int index) noexcept;
    IFACEMETHODIMP_(void) MoveWindowIntoZoneByDirection(_In_ HWND window, _In_ HWND zoneWindow, DWORD vkCode) noexcept;
    IFACEMETHODIMP_(void) MoveSizeExit(_In_ HWND window, _In_ HWND zoneWindow, POINT ptClient) noexcept;

private:
    void InitialPopulateZones() noexcept;
    void GenerateGridZones(_In_ MONITORINFO const& mi) noexcept;
    void DoGridLayout(SIZE const& zoneArea, int numCols, int numRows) noexcept;
    void GenerateFocusZones(_In_ MONITORINFO const& mi) noexcept;
    void StampZone(_In_ HWND window, _In_opt_ ComPtr<IZone> zone) noexcept;

    std::vector<ComPtr<IZone>> m_zones;
    ZoneSetConfig m_config;
};

IFACEMETHODIMP ZoneSet::AddZone(_In_ ComPtr<IZone> zone, bool front) noexcept
{
    // XXXX: need to reorder ids when inserting...
    if (front)
    {
        m_zones.insert(m_zones.begin(), zone);
    }
    else
    {
        m_zones.emplace_back(zone);
    }

    // Important not to set Id 0 since we store it in the HWND using SetProp.
    // SetProp(0) doesn't really work.
    zone->SetId(m_zones.size());
    return S_OK;
}

IFACEMETHODIMP ZoneSet::RemoveZone(_In_ ComPtr<IZone> zone) noexcept
{
    auto iter = std::find(m_zones.begin(), m_zones.end(), zone);
    if (iter != m_zones.end())
    {
        m_zones.erase(iter);
        return S_OK;
    }
    return E_INVALIDARG;
}

IFACEMETHODIMP_(ComPtr<IZone>) ZoneSet::ZoneFromPoint(POINT pt) noexcept
{
    for (auto iter = m_zones.begin(); iter != m_zones.end(); iter++)
    {
        ComPtr<IZone> zone;
        if (SUCCEEDED(iter->As(&zone)) && PtInRect(&zone->GetZoneRect(), pt))
        {
            return zone;
        }
    }
    return nullptr;
}

IFACEMETHODIMP_(ComPtr<IZone>) ZoneSet::ZoneFromWindow(_In_ HWND window) noexcept
{
    for (auto iter = m_zones.begin(); iter != m_zones.end(); iter++)
    {
        ComPtr<IZone> zone;
        if (SUCCEEDED(iter->As(&zone)) && zone->ContainsWindow(window))
        {
            return zone;
        }
    }
    return nullptr;
}

IFACEMETHODIMP_(ComPtr<IZoneSet>) ZoneSet::MakeCustomClone() noexcept
{
    if (SUCCEEDED(CoCreateGuid(&m_config.id)))
    {
        m_config.custom = true;
        return Make<ZoneSet>(m_config, m_zones);
    }
    return nullptr;
}

IFACEMETHODIMP_(void) ZoneSet::Save() noexcept
{
    size_t const zoneCount = m_zones.size();
    if (zoneCount == 0)
    {
        RegistryHelpers::DeleteZoneSet(m_config.workArea, m_config.id);
    }
    else
    {
        ZoneSetPersistedData data{};
        data.zoneCount = zoneCount;
        data.layout = m_config.layout;
        data.paddingInner = m_config.paddingInner;
        data.paddingOuter = m_config.paddingOuter;

        int i = 0;
        for (auto iter = m_zones.begin(); iter != m_zones.end(); iter++)
        {
            ComPtr<IZone> zone;
            if (SUCCEEDED(iter->As(&zone)))
            {
                CopyRect(&data.zones[i++], &zone->GetZoneRect());
            }
        }

        PWSTR guid;
        if (SUCCEEDED(StringFromCLSID(m_config.id, &guid)))
        {
            HKEY hkey = RegistryHelpers::CreateKey(m_config.workArea);
            if (hkey)
            {
                RegSetValueExW(hkey, guid, 0, REG_BINARY, reinterpret_cast<BYTE*>(&data), sizeof(data));
                RegCloseKey(hkey);
            }
            CoTaskMemFree(guid);
        }
    }
}

IFACEMETHODIMP_(void) ZoneSet::MoveZoneToFront(_In_ ComPtr<IZone> zone) noexcept
{
    auto iter = std::find(m_zones.begin(), m_zones.end(), zone);
    if (iter != m_zones.end())
    {
        std::rotate(m_zones.begin(), iter, iter + 1);
    }
}

IFACEMETHODIMP_(void) ZoneSet::MoveZoneToBack(_In_ ComPtr<IZone> zone) noexcept
{
    auto iter = std::find(m_zones.begin(), m_zones.end(), zone);
    if (iter != m_zones.end())
    {
        std::rotate(iter, iter + 1, m_zones.end());
    }
}

IFACEMETHODIMP_(void) ZoneSet::MoveWindowIntoZoneByIndex(_In_ HWND window, _In_ HWND windowZone, int index) noexcept
{
    if (index >= static_cast<int>(m_zones.size()))
    {
        index = 0;
    }

    auto zone = m_zones.at(index);
    if (zone)
    {
        zone->AddWindowToZone(window, windowZone, false);
    }
}

IFACEMETHODIMP_(void) ZoneSet::MoveWindowIntoZoneByDirection(_In_ HWND window, _In_ HWND windowZone, DWORD vkCode) noexcept
{
    ComPtr<IZone> oldZone;
    ComPtr<IZone> newZone;

    auto iter = std::find(m_zones.begin(), m_zones.end(), ZoneFromWindow(window));
    if (iter == m_zones.end())
    {
        iter = (vkCode == VK_RIGHT) ? m_zones.begin() : m_zones.end() - 1;
    }
    else if (SUCCEEDED(iter->As(&oldZone)))
    {
        if (vkCode == VK_LEFT)
        {
            if (iter == m_zones.begin())
            {
                iter = m_zones.end();
            }
            iter--;
        }
        else if (vkCode == VK_RIGHT)
        {
            iter++;
            if (iter == m_zones.end())
            {
                iter = m_zones.begin();
            }
        }
    }

    if (SUCCEEDED(iter->As(&newZone)))
    {
        if (oldZone)
        {
            oldZone->RemoveWindowFromZone(window, false);
        }
        newZone->AddWindowToZone(window, windowZone, true);
    }
}

IFACEMETHODIMP_(void) ZoneSet::MoveSizeExit(_In_ HWND window, _In_ HWND zoneWindow, POINT ptClient) noexcept
{
    auto zoneDrop = ZoneFromWindow(window);
    if (zoneDrop)
    {
        zoneDrop->RemoveWindowFromZone(window, !IsZoomed(window));
    }

    auto zone = ZoneFromPoint(ptClient);
    if (zone)
    {
        zone->AddWindowToZone(window, zoneWindow, true);

        POINT pointAdjustedScreen = ptClient;
        MapWindowPoints(zoneWindow, nullptr, &pointAdjustedScreen, 1);
        SetCursorPos(pointAdjustedScreen.x, pointAdjustedScreen.y);
    }
}

void ZoneSet::InitialPopulateZones() noexcept
{
    MONITORINFO mi{};
    mi.cbSize = sizeof(mi);
    if (GetMonitorInfoW(m_config.monitor, &mi))
    {
        // XXXX: does primary monitor mean anything? More likely to be 0,0?
        // XXXX: use monitor topology to infer monitors by direction. up is up

        if ((m_config.layout == ZoneSetLayout::Grid) || (m_config.layout == ZoneSetLayout::Row))
        {
            GenerateGridZones(mi);
        }
        else if (m_config.layout == ZoneSetLayout::Focus)
        {
            GenerateFocusZones(mi);
        }

        Save();
    }
}

void ZoneSet::GenerateGridZones(_In_ MONITORINFO const& mi) noexcept
{
    Rect workArea(mi.rcWork);

    int numCols, numRows;
    if (m_config.layout == ZoneSetLayout::Grid)
    {
        switch (m_config.zoneCount)
        {
            case 1: numCols = 1; numRows = 1; break;
            case 2: numCols = 2; numRows = 1; break;
            case 3: numCols = 2; numRows = 2; break;
            case 4: numCols = 2; numRows = 2; break;
            case 5: numCols = 3; numRows = 3; break;
            case 6: numCols = 3; numRows = 3; break;
            case 7: numCols = 3; numRows = 3; break;
            case 8: numCols = 3; numRows = 3; break;
            case 9: numCols = 3; numRows = 3; break;
        }

        if ((m_config.zoneCount == 2) && (workArea.height > workArea.width))
        {
            numCols = 1;
            numRows = 2;
        }
    }
    else if (m_config.layout == ZoneSetLayout::Row)
    {
        numCols = m_config.zoneCount;
        numRows = 1;
    }

    SIZE const zoneArea = {
        workArea.width - ((m_config.paddingOuter * 2) + (m_config.paddingInner * (numCols - 1))),
        workArea.height - ((m_config.paddingOuter * 2) + (m_config.paddingInner * (numRows - 1)))
    };

    DoGridLayout(zoneArea, numCols, numRows);
}

void ZoneSet::DoGridLayout(SIZE const& zoneArea, int numCols, int numRows) noexcept
{
    auto x = m_config.paddingOuter;
    auto y = m_config.paddingOuter;
    auto const zoneWidth = (zoneArea.cx / numCols);
    auto const zoneHeight = (zoneArea.cy / numRows);
    for (auto i = 1; i <= m_config.zoneCount; i++)
    {
        auto col = numCols - (i % numCols);
        RECT const zoneRect = { x, y, x + zoneWidth, y + zoneHeight };
        AddZone(MakeZone(zoneRect).Get(), false);

        x += zoneWidth + m_config.paddingInner;
        if (col == numCols)
        {
            x = m_config.paddingOuter;
            y += zoneHeight + m_config.paddingInner;
        }
    }
}

void ZoneSet::GenerateFocusZones(_In_ MONITORINFO const& mi) noexcept
{
    Rect const workArea(mi.rcWork);

    SIZE const workHalf = { workArea.width / 2, workArea.height / 2 };
    RECT const safeZone = {
        m_config.paddingOuter,
        m_config.paddingOuter,
        workArea.width - m_config.paddingOuter,
        workArea.height - m_config.paddingOuter
    };

    int const width = min(1920, workArea.width * 60 / 100);
    int const height = min(1200, workArea.height * 75 / 100);
    int const halfWidth = width / 2;
    int const halfHeight = height / 2;
    int x = workHalf.cx - halfWidth;
    int y = workHalf.cy - halfHeight;

    RECT const focusRect = { x, y, x + width, y + height };
    AddZone(MakeZone(focusRect).Get(), false);

    for (auto i = 2; i <= m_config.zoneCount; i++)
    {
        switch (i)
        {
            case 2: x = focusRect.right - halfWidth; y = focusRect.top + m_config.paddingInner; break; // right
            case 3: x = focusRect.left - halfWidth; y = focusRect.top + (m_config.paddingInner * 2); break; // left
            case 4: x = focusRect.left + m_config.paddingInner; y = focusRect.top - halfHeight; break; // up
            case 5: x = focusRect.left - m_config.paddingInner; y = focusRect.bottom - halfHeight; break; // down
        }

        // Bound into safe zone
        x = min(safeZone.right - width, max(safeZone.left, x));
        y = min(safeZone.bottom - height, max(safeZone.top, y));

        RECT const zoneRect = { x, y, x + width, y + height };
        AddZone(MakeZone(zoneRect).Get(), false);
    }
}

ComPtr<IZoneSet> MakeZoneSet(ZoneSetConfig const& config) noexcept
{
    return Make<ZoneSet>(config);
}