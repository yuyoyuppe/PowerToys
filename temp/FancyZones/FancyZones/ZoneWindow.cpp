#include "stdafx.h"
#include <ShellScalingApi.h>

using namespace Microsoft::WRL;

class ZoneWindow WrlFinal : public RuntimeClass<
    RuntimeClassFlags<Microsoft::WRL::ClassicCom>,
    IZoneWindow>
{
public:
    ZoneWindow(_In_ HMONITOR monitor, _In_ PCWSTR deviceId);

    ~ZoneWindow()
    {
        DestroyWindow(m_window);
        CoTaskMemFree(m_deviceId);
    }

    IFACEMETHODIMP ShowZoneWindow(bool activate) noexcept;
    IFACEMETHODIMP HideZoneWindow() noexcept;
    IFACEMETHODIMP MoveSizeEnter(_In_ HWND window, POINT ptScreen, DragMode dragMode) noexcept;
    IFACEMETHODIMP MoveSizeExit(_In_ HWND window, POINT ptScreen) noexcept;
    IFACEMETHODIMP MoveSizeUpdate(POINT ptScreen, DragMode dragMode) noexcept;
    IFACEMETHODIMP MoveSizeCancel() noexcept;
    IFACEMETHODIMP_(DragMode) GetDragMode() noexcept { return m_dragMode; }
    IFACEMETHODIMP_(void) MoveWindowIntoZoneByIndex(_In_ HWND window, int index) noexcept;
    IFACEMETHODIMP_(void) MoveWindowIntoZoneByDirection(_In_ HWND window, DWORD vkCode) noexcept;
    IFACEMETHODIMP_(void) OnDisplayChange(DisplayChangeType type) noexcept;
    IFACEMETHODIMP_(void) CycleActiveZoneSet(DWORD vkCode) noexcept;

protected:
    static LRESULT CALLBACK s_WndProc(HWND window, UINT message, WPARAM wparam, LPARAM lparam) noexcept;

private:
    struct ColorSetting
    {
        BYTE fillAlpha{};
        COLORREF fill{};
        BYTE borderAlpha{};
        COLORREF border{};
        int thickness{};
    };

    void InitializeId(PCWSTR deviceId) noexcept;
    void LoadSettings() noexcept;
    void InitializeZoneSets() noexcept;
    void LoadZoneSetsFromRegistry() noexcept;
    ComPtr<IZoneSet> AddZoneSet(ZoneSetLayout layout, int numZones, int paddingOuter, int paddingInner) noexcept;
    void MakeActiveZoneSetCustom() noexcept;
    void UpdateActiveZoneSet(_In_opt_ IZoneSet* zoneSet) noexcept;
    LRESULT WndProc(UINT message, WPARAM wparam, LPARAM lparam) noexcept;
    void OnTimer(WPARAM wparam) noexcept;
    void OnLButtonDown(_In_ LPARAM lparam) noexcept;
    void OnLButtonUp(_In_ LPARAM lparam) noexcept;
    void OnRButtonUp(_In_ LPARAM lparam) noexcept;
    void OnMouseMove(_In_ LPARAM lparam) noexcept;
    void DrawBackdrop(_In_ HDC hdc, RECT const& clientRect) noexcept;
    void DrawGridLines(_In_ HDC hdc, RECT const& clientRect) noexcept;
    void DrawZone(_In_ HDC hdc, ColorSetting const& colorSetting, ComPtr<IZone> zone) noexcept;
    void DrawIndex(_In_ HDC hdc, POINT offset, int index, int padding, int size, bool flipX, bool flipY, COLORREF colorFill);
    void DrawActiveZoneSet(_In_ HDC hdc, RECT const& clientRect) noexcept;
    void DrawZoneBuilder(_In_ HDC hdc, RECT const& clientRect) noexcept;
    void DrawSwitchButtons(_In_ HDC hdc, RECT const& clientRect) noexcept;
    void OnPaint(_In_ HDC hdc) noexcept;
    void UpdateGrid(int stepColumns, int stepRows) noexcept;
    void UpdateGridMargins(int inc) noexcept;
    void EnterEditorMode() noexcept;
    void ExitEditorMode() noexcept;
    void OnKeyUp(_In_ WPARAM wparam) noexcept;
    ComPtr<IZone> ZoneFromPoint(POINT pt) noexcept;
    void ChooseDefaultActiveZoneSet() noexcept;
    bool IsOccluded(POINT pt, size_t index) noexcept;
    void CycleActiveZoneSetInternal(DWORD wparam) noexcept;
    void FlashZones(bool delayed) noexcept;
    int GetSwitchButtonIndexFromPoint(POINT ptClient) noexcept;
    UINT GetDpiForMonitor() noexcept;

    HMONITOR m_monitor{};
    wchar_t m_uniqueId[256]{};  // Parsed deviceId + resolution + virtualDesktopId
    wchar_t m_workArea[256]{};
    PWSTR m_deviceId{};
    HWND m_window{};
    HWND m_windowMoveSize{};
    bool m_buttonDown{};
    bool m_drawHints{};
    bool m_editorMode{};
    bool m_flashMode{};
    POINT m_ptDown{};
    POINT m_ptLast{};
    POINT m_moveSizeEnterClient{};
    POINT m_pointAdjustedClient{};
    DragMode m_dragMode = DragMode::None;
    ComPtr<IZoneSet> m_activeZoneSet;
    GUID m_activeZoneSetId{};
    std::vector<ComPtr<IZoneSet>> m_zoneSets;
    ComPtr<IZone> m_highlightZone;
    WPARAM m_keyLast{};
    size_t m_keyCycle{};
    int m_gridWidth{};
    int m_gridHeight{};
    int m_gridRows{};
    int m_gridColumns{};
    int m_flashTimer{};
    int m_switchButtonWidth = 50;
    int m_switchButtonPadding = 5;
    int m_switchButtonHover = -1;
    SIZE m_gridMargins{};
    RECT m_zoneBuilder{};
    RECT m_switchButtonContainerRect{};

    static const UINT m_showAnimationDuration = 200;
};

ZoneWindow::ZoneWindow(_In_ HMONITOR monitor, _In_ PCWSTR deviceId) :
    m_monitor(monitor)
{
    MONITORINFO mi{};
    mi.cbSize = sizeof(mi);
    if (GetMonitorInfoW(m_monitor, &mi))
    {
        auto dpi = GetDpiForMonitor();

        Rect workArea(mi.rcWork);
        workArea.right = workArea.left + MulDiv(workArea.width, dpi, 96);
        workArea.bottom = workArea.top + MulDiv(workArea.height, dpi, 96);

        StringCchPrintf(m_workArea, ARRAYSIZE(m_workArea), L"%d_%d", workArea.width, workArea.height);

        InitializeId(deviceId);
        LoadSettings();
        InitializeZoneSets();

        WNDCLASSEXW wcex{};
        wcex.cbSize = sizeof(WNDCLASSEX);
        wcex.lpfnWndProc = s_WndProc;
        wcex.hInstance = g_app->GetHInstance();
        wcex.lpszClassName = L"SuperFancyZones_ZoneWindow";
        wcex.hCursor = LoadCursorW(nullptr, IDC_ARROW);
        RegisterClassExW(&wcex);

        m_window = CreateWindowExW(WS_EX_TOOLWINDOW, L"SuperFancyZones_ZoneWindow", L"", WS_POPUP, workArea.left, workArea.top, workArea.width, workArea.height, nullptr, nullptr, g_app->GetHInstance(), this);
        if (m_window)
        {
            MakeWindowTransparent(m_window);
            UpdateGrid(0, 0);
            if (g_app->GetSettings().virtualDesktopChange_flashZones)
            {
                FlashZones(true /*delayed*/);
            }
        }
    }
}

IFACEMETHODIMP ZoneWindow::ShowZoneWindow(bool activate) noexcept
{
    if (m_window)
    {
        m_flashMode = false;

        UINT flags = SWP_NOSIZE | SWP_NOMOVE;
        if (!activate)
        {
            flags |= SWP_NOACTIVATE;
        }

        HWND windowInsertAfter = m_windowMoveSize;
        if (windowInsertAfter == nullptr)
        {
            windowInsertAfter = HWND_TOPMOST;
        }
        SetWindowPos(m_window, windowInsertAfter, 0, 0, 0, 0, flags);

        AnimateWindow(m_window, m_showAnimationDuration, AW_BLEND);

        return S_OK;
    }
    return E_FAIL;
}

IFACEMETHODIMP ZoneWindow::HideZoneWindow() noexcept
{
    if (!m_window)
    {
        return E_FAIL;
    }

    ShowWindow(m_window, SW_HIDE);
    m_keyLast = 0;
    m_windowMoveSize = nullptr;
    m_drawHints = false;
    m_highlightZone = nullptr;
    m_editorMode = false;
    return S_OK;
}

IFACEMETHODIMP ZoneWindow::MoveSizeEnter(_In_ HWND window, POINT ptScreen, DragMode dragMode) noexcept
{
    // XXXX: could forward input to a background thread
    // XXXX: then we wouldn't have to disrupt the UI thread

    if (m_windowMoveSize)
    {
        return E_INVALIDARG;
    }

    m_moveSizeEnterClient = ptScreen;
    MapWindowPoints(nullptr, m_window, &m_moveSizeEnterClient, 1);
    m_dragMode = dragMode;
    m_windowMoveSize = window;
    m_drawHints = true;
    m_highlightZone = nullptr;
    ShowZoneWindow(false);
    return S_OK;
}

IFACEMETHODIMP ZoneWindow::MoveSizeExit(_In_ HWND window, POINT ptScreen) noexcept
{
    if (m_windowMoveSize != window)
    {
        return E_INVALIDARG;
    }

    if (m_activeZoneSet)
    {
        m_activeZoneSet->MoveSizeExit(window, m_window, m_pointAdjustedClient);
    }

    HideZoneWindow();
    m_windowMoveSize = nullptr;
    return S_OK;
}

IFACEMETHODIMP ZoneWindow::MoveSizeUpdate(POINT ptScreen, DragMode dragMode) noexcept
{
    bool redraw = false;
    POINT ptClient = ptScreen;
    MapWindowPoints(nullptr, m_window, &ptClient, 1);

    if (dragMode != m_dragMode)
    {
        m_dragMode = dragMode;
        m_moveSizeEnterClient = ptClient;
    }

    if (m_dragMode == DragMode::Normal)
    {
        auto highlight = ZoneFromPoint(ptClient);
        redraw = (highlight != m_highlightZone);
        m_highlightZone = std::move(highlight);
        m_pointAdjustedClient = ptClient;
    }
    else if (m_dragMode == DragMode::SetWindowPos)
    {
    }
    else
    {
        return E_INVALIDARG;
    }

    if (redraw)
    {
        InvalidateRect(m_window, nullptr, true);
    }
    return S_OK;
}

IFACEMETHODIMP ZoneWindow::MoveSizeCancel() noexcept
{
    HideZoneWindow();
    return S_OK;
}

IFACEMETHODIMP_(void) ZoneWindow::MoveWindowIntoZoneByIndex(_In_ HWND window, int index) noexcept
{
    if (m_activeZoneSet)
    {
        m_activeZoneSet->MoveWindowIntoZoneByIndex(window, m_window, index);
    }
}

IFACEMETHODIMP_(void) ZoneWindow::MoveWindowIntoZoneByDirection(_In_ HWND window, DWORD vkCode) noexcept
{
    if (m_activeZoneSet)
    {
        m_activeZoneSet->MoveWindowIntoZoneByDirection(window, m_window, vkCode);
    }
}

IFACEMETHODIMP_(void) ZoneWindow::OnDisplayChange(DisplayChangeType type) noexcept
{
    // We're about to be destroyed.
}

IFACEMETHODIMP_(void) ZoneWindow::CycleActiveZoneSet(DWORD wparam) noexcept
{
    CycleActiveZoneSetInternal(wparam);

    if (m_windowMoveSize)
    {
        InvalidateRect(m_window, nullptr, true);
    }
    else if (g_app->GetSettings().zoneSetChange_flashZones)
    {
        FlashZones(false /*delayed*/);
    }
}

#pragma region private
void ParseDeviceId(PCWSTR deviceId, PWSTR parsedId, size_t size)
{
    // We're interested in the unique part between the first and last #'s
    // Example input: \\?\DISPLAY#DELA026#5&10a58c63&0&UID16777488#{e6f07b5f-ee97-4a90-b076-33f57bf4eaa7}
    // Example output: DELA026#5&10a58c63&0&UID16777488 

    wchar_t buffer[256];
    StringCchCopy(buffer, 256, deviceId);

    PWSTR pszStart = wcschr(buffer, L'#');
    PWSTR pszEnd = wcsrchr(buffer, L'#');
    if (pszStart && pszEnd && (pszStart != pszEnd))
    {
        pszStart++; // skip past the first #
        *pszEnd = '\0';
        StringCchCopy(parsedId, size, pszStart);
    }
    else
    {
        StringCchCopy(parsedId, size, L"FallbackDevice");
    }
}

void ZoneWindow::InitializeId(PCWSTR deviceId) noexcept
{
    SHStrDup(deviceId, &m_deviceId);

    MONITORINFOEXW mi;
    mi.cbSize = sizeof(mi);
    if (GetMonitorInfo(m_monitor, &mi))
    {
        GUID const currentVirtualDesktopId = g_app->GetCurrentVirtualDesktopId();
        PWSTR virtualDesktop;
        if (SUCCEEDED(StringFromCLSID(currentVirtualDesktopId, &virtualDesktop)))
        {
            wchar_t parsedId[256]{};
            ParseDeviceId(m_deviceId, parsedId, 256);

            Rect const monitorRect(mi.rcWork);
            StringCchPrintf(m_uniqueId, ARRAYSIZE(m_uniqueId), L"%s_%d_%d_%s",
                parsedId, monitorRect.width, monitorRect.height, virtualDesktop);
        }
    }
}

void ZoneWindow::LoadSettings() noexcept
{
    RegistryHelpers::GetValue<GUID>(m_uniqueId, L"ActiveZoneSetId", &m_activeZoneSetId, sizeof(m_activeZoneSetId));
    RegistryHelpers::GetValue<SIZE>(m_uniqueId, L"GridMargins", &m_gridMargins, sizeof(m_gridMargins));
}

void ZoneWindow::InitializeZoneSets() noexcept
{
    // XXXX: does primary monitor mean anything? More likely to be 0,0?
    // XXXX: use monitor topology to infer monitors by direction. up is up
    // XXXX: make setting to generate zones with padding

    LoadZoneSetsFromRegistry();
    if (m_zoneSets.empty())
    {
        int const paddingOuter = 40;
        int const paddingInner = 20;

        for (int numZones = 2; numZones <= 5; numZones++)
        {
            AddZoneSet(ZoneSetLayout::Focus, numZones, paddingOuter, paddingInner);
        }

        for (int numZones = 1; numZones <= 9; numZones++)
        {
            AddZoneSet(ZoneSetLayout::Grid, numZones, paddingOuter, paddingInner);
            AddZoneSet(ZoneSetLayout::Grid, numZones, 0, 0);
        }

        for (int numZones = 3; numZones <= 6; numZones++)
        {
            AddZoneSet(ZoneSetLayout::Row, numZones, paddingOuter, paddingInner);
            AddZoneSet(ZoneSetLayout::Row, numZones, 0, 0);
        }
    }

    if (!m_activeZoneSet)
    {
        ChooseDefaultActiveZoneSet();
    }
}

void ZoneWindow::LoadZoneSetsFromRegistry() noexcept
{
    HKEY key = RegistryHelpers::OpenKey(m_workArea);
    if (key)
    {
        ZoneSetPersistedData data{};
        DWORD dataSize = sizeof(data);
        wchar_t value[256]{};
        DWORD valueLength = ARRAYSIZE(value);
        DWORD i = 0;
        while (RegEnumValueW(key, i++, value, &valueLength, nullptr, nullptr, reinterpret_cast<BYTE*>(&data), &dataSize) == ERROR_SUCCESS)
        {
            if (data.version == VERSION_PERSISTEDDATA)
            {
                GUID zoneSetId;
                if (SUCCEEDED(CLSIDFromString(value, &zoneSetId)))
                {
                    auto zoneSet = MakeZoneSet(ZoneSetConfig(
                        zoneSetId,
                        m_monitor,
                        m_workArea,
                        data.layout,
                        0,
                        static_cast<int>(data.paddingInner),
                        static_cast<int>(data.paddingOuter)));
                    if (zoneSet)
                    {
                        for (UINT j = 0; j < data.zoneCount; j++)
                        {
                            zoneSet->AddZone(MakeZone(data.zones[j]), false);
                        }

                        m_zoneSets.emplace_back(zoneSet);

                        if (zoneSetId == m_activeZoneSetId)
                        {
                            UpdateActiveZoneSet(zoneSet.Get());
                        }
                    }
                }
            }
            else
            {
                // XXXX: Setting migration
            }

            valueLength = ARRAYSIZE(value);
            dataSize = sizeof(data);
        }
        RegCloseKey(key);
    }
}

ComPtr<IZoneSet> ZoneWindow::AddZoneSet(ZoneSetLayout layout, int numZones, int paddingOuter, int paddingInner) noexcept
{
    GUID zoneSetId;
    if (SUCCEEDED(CoCreateGuid(&zoneSetId)))
    {
        auto zoneSet = MakeZoneSet(ZoneSetConfig(zoneSetId, m_monitor, m_workArea, layout, numZones, paddingOuter, paddingInner));
        if (zoneSet)
        {
            m_zoneSets.emplace_back(zoneSet);
            return zoneSet;
        }
    }
    return nullptr;
}

void ZoneWindow::MakeActiveZoneSetCustom() noexcept
{
    if (m_activeZoneSet)
    {
        auto customZoneSet = m_activeZoneSet->MakeCustomClone();
        if (customZoneSet)
        {
            UpdateActiveZoneSet(customZoneSet.Get());
            m_zoneSets.emplace_back(customZoneSet);
        }
    }
}

void ZoneWindow::UpdateActiveZoneSet(_In_opt_ IZoneSet* zoneSet) noexcept
{
    m_activeZoneSet = zoneSet;

    if (m_activeZoneSet)
    {
        RegistryHelpers::SetValue<GUID>(m_uniqueId, L"ActiveZoneSetId", m_activeZoneSet->GetId(), sizeof(GUID));
    }
}

LRESULT ZoneWindow::WndProc(UINT message, WPARAM wparam, LPARAM lparam) noexcept
{
    switch (message)
    {
        case WM_NCDESTROY:
        {
            ::DefWindowProc(m_window, message, wparam, lparam);
            SetWindowLongPtr(m_window, GWLP_USERDATA, 0);
        }
        break;

        case WM_ERASEBKGND:
            return 1;

        case WM_PRINTCLIENT:
        case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = reinterpret_cast<HDC>(wparam);
            if (!hdc)
            {
                hdc = BeginPaint(m_window, &ps);
            }

            OnPaint(hdc);

            if (wparam == 0)
            {
                EndPaint(m_window, &ps);
            }
        }
        break;

        case WM_TIMER:
            OnTimer(wparam);
            break;

        case WM_LBUTTONDOWN:
            OnLButtonDown(lparam);
            break;

        case WM_LBUTTONUP:
            OnLButtonUp(lparam);
            break;

        case WM_RBUTTONUP:
            OnRButtonUp(lparam);
            break;

        case WM_MOUSEMOVE:
            OnMouseMove(lparam);
            break;

        case WM_KEYUP:
            OnKeyUp(wparam);
            break;

        default:
        {
            return DefWindowProc(m_window, message, wparam, lparam);
        }
    }
    return 0;
}

void ZoneWindow::OnTimer(WPARAM wparam) noexcept
{
    if (wparam == 1)
    {
        ShowWindow(m_window, SW_SHOWNA);
        KillTimer(m_window, m_flashTimer);
        AnimateWindow(m_window, 700, AW_HIDE | AW_BLEND);
    }
}

void ZoneWindow::OnLButtonDown(_In_ LPARAM lparam) noexcept
{
    m_buttonDown = true;
    m_ptDown = { GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam) };
}

void ZoneWindow::OnLButtonUp(_In_ LPARAM lparam) noexcept
{
    POINT const ptClient = { GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam) };
    if (m_buttonDown && m_activeZoneSet)
    {
        if (m_editorMode)
        {
            bool const ctrl = GetAsyncKeyState(VK_CONTROL) & 0x8000;
            if (ctrl)
            {
                auto zone = ZoneFromPoint(ptClient);
                if (zone)
                {
                    m_activeZoneSet->RemoveZone(zone);

                    int const padding = m_activeZoneSet->GetInnerPadding();
                    RECT const zoneRect = zone->GetZoneRect();
                    int const zoneRectWidthHalf = ((zoneRect.right - zoneRect.left) / 2) - padding;
                    RECT rectLeft = zoneRect;
                    rectLeft.right = rectLeft.left + zoneRectWidthHalf;
                    m_activeZoneSet->AddZone(MakeZone(rectLeft), false);

                    RECT rectRight = zoneRect;
                    rectRight.left = rectLeft.right + padding;
                    m_activeZoneSet->AddZone(MakeZone(rectRight), false);

                    m_activeZoneSet->Save();
                }
            }
            else if (m_activeZoneSet && !IsRectEmpty(&m_zoneBuilder))
            {
                m_activeZoneSet->AddZone(MakeZone(m_zoneBuilder), true);
            }
        }
        else if (!m_flashMode && !m_editorMode && !m_drawHints)
        {
            if (PtInRect(&m_switchButtonContainerRect, ptClient))
            {
                auto switchButtonIndex = GetSwitchButtonIndexFromPoint(ptClient);
                if (switchButtonIndex != -1)
                {
                    CycleActiveZoneSetInternal('0' + switchButtonIndex);
                }
            }
            else
            {
                auto zone = ZoneFromPoint(ptClient);
                if (zone)
                {
                    m_activeZoneSet->MoveZoneToFront(zone);
                    m_activeZoneSet->Save();
                }
            }
        }
    }

    m_zoneBuilder = {};
    m_buttonDown = false;
    InvalidateRect(m_window, nullptr, true);
}

void ZoneWindow::OnRButtonUp(_In_ LPARAM lparam) noexcept
{
    POINT const ptClient = { GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam) };
    if (m_activeZoneSet)
    {
        if (m_editorMode)
        {
            bool const ctrl = GetAsyncKeyState(VK_CONTROL) & 0x8000;
            if (ctrl)
            {
                auto zone = ZoneFromPoint(ptClient);
                if (zone)
                {
                    m_activeZoneSet->RemoveZone(zone);

                    int const padding = m_activeZoneSet->GetInnerPadding();
                    RECT const zoneRect = zone->GetZoneRect();
                    int const zoneRectHeightHalf = ((zoneRect.bottom - zoneRect.top) / 2) - padding;
                    RECT rectTop = zoneRect;
                    rectTop.bottom = rectTop.top + zoneRectHeightHalf;
                    m_activeZoneSet->AddZone(MakeZone(rectTop), false);

                    RECT rectBottom = zoneRect;
                    rectBottom.top = rectTop.bottom + padding;
                    m_activeZoneSet->AddZone(MakeZone(rectBottom), false);

                    m_activeZoneSet->Save();
                }
            }
            else
            {
                auto zone = ZoneFromPoint(ptClient);
                if (zone)
                {
                    m_activeZoneSet->RemoveZone(zone);
                    m_activeZoneSet->Save();
                }
            }
        }
        else
        {
                auto zone = ZoneFromPoint(ptClient);
                if (zone)
                {
                    m_activeZoneSet->MoveZoneToBack(zone);
                    m_activeZoneSet->Save();
                }
        }
    }
    InvalidateRect(m_window, nullptr, true);
}

void ZoneWindow::OnMouseMove(_In_ LPARAM lparam) noexcept
{
    int const oldHover = m_switchButtonHover;
    m_switchButtonHover = -1;

    POINT const ptClient = { GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam) };
    if (m_buttonDown && m_editorMode && ((ptClient.x != m_ptLast.x) || (ptClient.y != m_ptLast.y)))
    {
        // XXXX: this is gross
        RECT start;
        int const startColumn = max(0, min(m_gridColumns, ((m_ptDown.x - m_gridMargins.cx) / m_gridWidth)));
        int const startRow = max(0, min(m_gridRows, ((m_ptDown.y  - m_gridMargins.cy) / m_gridHeight)));
        start.left = startColumn * m_gridWidth;
        start.top = startRow * m_gridHeight;
        start.right = start.left + m_gridWidth;
        start.bottom = start.top + m_gridHeight;
        OffsetRect(&start, m_gridMargins.cx, m_gridMargins.cy);

        RECT current;
        int const currentColumn = max(0, min(m_gridColumns, ((ptClient.x - m_gridMargins.cx) / m_gridWidth)));
        int const currentRow = max(0, min(m_gridRows, ((ptClient.y - m_gridMargins.cy) / m_gridHeight)));
        current.left = currentColumn * m_gridWidth;
        current.top = currentRow * m_gridHeight;
        current.right = current.left + m_gridWidth;
        current.bottom = current.top + m_gridHeight;
        OffsetRect(&current, m_gridMargins.cx, m_gridMargins.cy);

        RECT invalidateRect = m_zoneBuilder;
        POINT const last = {
             m_gridMargins.cx + ((m_ptLast.x - m_gridMargins.cx) / m_gridWidth),
             m_gridMargins.cy + ((m_ptLast.y - m_gridMargins.cy) / m_gridHeight)
        };

        m_ptLast = ptClient;
        UnionRect(&m_zoneBuilder, &start, &current);
        UnionRect(&invalidateRect, &invalidateRect, &m_zoneBuilder);

        if ((current.left != last.x) || (current.top != last.y))
        {
            InvalidateRect(m_window, &invalidateRect, true);
        }
    }
    else if (!m_flashMode && !m_editorMode && !m_drawHints && PtInRect(&m_switchButtonContainerRect, ptClient))
    {
        m_switchButtonHover = GetSwitchButtonIndexFromPoint(ptClient);
    }

    if (oldHover != m_switchButtonHover)
    {
        InvalidateRect(m_window, &m_switchButtonContainerRect, true);
    }
}

void ZoneWindow::DrawBackdrop(_In_ HDC hdc, RECT const& clientRect) noexcept
{
    if (m_windowMoveSize || m_flashMode)
    {
        FillRectARGB(hdc, &clientRect, 0, RGB(0, 0, 0), false);
    }
    else
    {
        FillRectARGB(hdc, &clientRect, 225, RGB(0, 0, 0), false);
    }
}

void ZoneWindow::DrawGridLines(_In_ HDC hdc, RECT const& clientRect) noexcept
{
    if (m_editorMode)
    {
        COLORREF const color = RGB(225, 225, 225);

        HPEN pen = CreatePen(PS_SOLID, 1, color);
        auto oldPen = SelectObject(hdc, pen);
        for (int i = 0; i <= m_gridRows; i++)
        {
            int const y = m_gridMargins.cy + (i * m_gridHeight);
            MoveToEx(hdc, m_gridMargins.cx, y, nullptr);
            LineTo(hdc, clientRect.right - m_gridMargins.cx, y);
        }

        for (int i = 0; i <= m_gridColumns; i++)
        {
            int const x = m_gridMargins.cx + (i * m_gridWidth);
            MoveToEx(hdc, x, m_gridMargins.cy, nullptr);
            LineTo(hdc, x, clientRect.bottom - m_gridMargins.cy);
        }
        SelectObject(hdc, oldPen);
        DeletePen(pen);
    }
}

void ZoneWindow::DrawZone(_In_ HDC hdc, ColorSetting const& colorSetting, ComPtr<IZone> zone) noexcept
{
    RECT zoneRect = zone->GetZoneRect();
    if (colorSetting.borderAlpha > 0)
    {
        FillRectARGB(hdc, &zoneRect, colorSetting.borderAlpha, colorSetting.border, false);
        InflateRect(&zoneRect, colorSetting.thickness, colorSetting.thickness);
    }
    FillRectARGB(hdc, &zoneRect, colorSetting.fillAlpha, colorSetting.fill, false);

    if (!m_flashMode)
    {
        COLORREF const colorFill = RGB(255, 255, 255);

        auto const index = zone->GetId();
        int const padding = 5;
        int const size = 10;
        POINT offset = { zoneRect.left + padding, zoneRect.top + padding };
        if (!IsOccluded(offset, index))
        {
            DrawIndex(hdc, offset, index, padding, size, false, false, colorFill); // top left
            return;
        }

        offset.x = zoneRect.right - ((padding + size) * 3);
        if (!IsOccluded(offset, index))
        {
            DrawIndex(hdc, offset, index, padding, size, true, false, colorFill); // top right
            return;
        }

        offset.y = zoneRect.bottom - ((padding + size) * 3);
        if (!IsOccluded(offset, index))
        {
            DrawIndex(hdc, offset, index, padding, size, true, true, colorFill); // bottom right
            return;
        }

        offset.x = zoneRect.left + padding;
        DrawIndex(hdc, offset, index, padding, size, false, true, colorFill); // bottom left
    }
}

void ZoneWindow::DrawIndex(_In_ HDC hdc, POINT offset, int index, int padding, int size, bool flipX, bool flipY, COLORREF colorFill)
{
    RECT rect = { offset.x, offset.y, offset.x + size, offset.y + size };
    for (int y = 0; y < 3; y++)
    {
        for (int x = 0; x < 3; x++)
        {
            if (index-- > 0)
            {
                RECT useRect = rect;
                if (flipX)
                {
                    if (x == 0) useRect.left += (size + padding + size + padding);
                    else if (x == 2) useRect.left -= (size + padding + size + padding);
                    useRect.right = useRect.left + size;
                }

                if (flipY)
                {
                    if (y == 0) useRect.top += (size + padding + size + padding);
                    else if (y == 2) useRect.top -= (size + padding + size + padding);
                    useRect.bottom = useRect.top + size;
                }

                FillRectARGB(hdc, &useRect, 200, RGB(50, 50, 50), true);

                RECT inside = useRect;
                InflateRect(&inside, -2, -2);

                FillRectARGB(hdc, &inside, 100, colorFill, true);

                rect.left += (size + padding);
                rect.right = rect.left + size;
            }
        }
        rect.left = offset.x;
        rect.right = rect.left + size;
        rect.top += (size + padding);
        rect.bottom = rect.top + size;
    }
}

void ZoneWindow::DrawActiveZoneSet(_In_ HDC hdc, RECT const& clientRect) noexcept
{
    if (m_activeZoneSet)
    {
        static COLORREF const grayColors[] = {
            RGB(75, 75, 85),
            RGB(150, 150, 160),
            RGB(100, 100, 110),
            RGB(125, 125, 135),
            RGB(225, 225, 235),
            RGB(25, 25, 35),
            RGB(200, 200, 210),
            RGB(50, 50, 60),
            RGB(175, 175, 185),
        };

        static COLORREF const brightColors[] = {
            RGB(0, 183, 195),
            RGB(255, 140, 0),
            RGB(1, 133, 116),
            RGB(231, 72, 86),
            RGB(191, 0, 119),
            RGB(116, 77, 169),
            RGB(0, 120, 215),
            RGB(45, 125, 154),
            RGB(86, 124, 115),
        };

        const COLORREF* colors = g_app->GetSettings().colorfulZones ? brightColors : grayColors;

        ColorSetting const colorHints      { 225, RGB(81, 92, 107),   255, RGB(104, 118, 138), -2 };
        ColorSetting const colorEditorMode { 240, RGB(100, 100, 100), 255, RGB(50, 50, 50),    -5 };
        ColorSetting       colorViewer     { 225, 0, /*colors*/       255, RGB(40, 50, 60),    -2 };
        ColorSetting const colorHighlight  { 225, RGB(0, 120, 215),   255, RGB(0, 99, 177),    -2 };
        ColorSetting const colorFlash      { 200, RGB(81, 92, 107),   200, RGB(104, 118, 138), -2 };

        auto zones = m_activeZoneSet->GetZones();
        int colorIndex = zones.size() - 1;
        for (auto iter = zones.rbegin(); iter != zones.rend(); iter++)
        {
            ComPtr<IZone> zone;
            if (SUCCEEDED(iter->As(&zone)))
            {
                if (zone != m_highlightZone)
                {
                    if (m_flashMode)
                    {
                        DrawZone(hdc, colorFlash, zone);
                    }
                    else if (m_drawHints)
                    {
                        DrawZone(hdc, colorHints, zone);
                    }
                    else if (m_editorMode)
                    {
                        DrawZone(hdc, colorEditorMode, zone);
                    }
                    else
                    {
                        colorViewer.fill = colors[colorIndex];
                        DrawZone(hdc, colorViewer, zone);
                    }
                }
                colorIndex--;
            }
        }

        if (m_highlightZone)
        {
            // XXXX: get the accent color
            DrawZone(hdc, colorHighlight, m_highlightZone);
        }
    }
}

void ZoneWindow::DrawZoneBuilder(_In_ HDC hdc, RECT const& clientRect) noexcept
{
    if (m_editorMode && m_buttonDown)
    {
        COLORREF const colorDrag = RGB(255, 255, 255);
        FillRectARGB(hdc, &m_zoneBuilder, 255, colorDrag, false);
    }
}

void ZoneWindow::DrawSwitchButtons(_In_ HDC hdc, RECT const& clientRect) noexcept
{
    if (!m_editorMode && !m_drawHints && !m_flashMode)
    {
        Rect const rect(clientRect);

        int const numButtons = 9;
        int const containerRectWidth = (m_switchButtonWidth * numButtons) + (m_switchButtonPadding * numButtons) + m_switchButtonPadding;
        int const containerRectHeight = 42;

        m_switchButtonContainerRect = { (rect.width / 2) - (containerRectWidth / 2), 0, (rect.width / 2) + (containerRectWidth / 2), containerRectHeight };

        COLORREF const switchButtonContainerColor = RGB(50, 50, 50);
        BYTE const switchButtonContainerAlpha = 150;
        FillRectARGB(hdc, &m_switchButtonContainerRect, switchButtonContainerAlpha, switchButtonContainerColor, true);

        COLORREF const fillColor = RGB(128, 128, 128);
        COLORREF const hoverColor = RGB(255, 255, 255);
        COLORREF const activeColor = RGB(0, 128, 0);
        COLORREF const activeHoverColor = RGB(0, 200, 0);

        size_t activeZoneCount = 0;
        if (m_activeZoneSet)
        {
            activeZoneCount = m_activeZoneSet->GetZones().size();
        }

        int x = m_switchButtonContainerRect.left + m_switchButtonPadding;
        for (UINT i = 1; i < 10; i++)
        {
            POINT const offset = { x, 5 };
            int const padding = 1;
            int const size = 10;

            bool const active = activeZoneCount == i;
            bool const hover = i == m_switchButtonHover;

            COLORREF color = fillColor;
            if (active && hover)
            {
                color = activeHoverColor;
            }
            else if (active)
            {
                color = activeColor;
            }
            else if (hover)
            {
                color = hoverColor;
            }

            DrawIndex(hdc, offset, i, padding, size, false /*flipX*/, false /*flipY*/, color);
            x += m_switchButtonWidth + m_switchButtonPadding;
        }
    }
}

void ZoneWindow::OnPaint(_In_ HDC hdc) noexcept
{
    RECT clientRect;
    GetClientRect(m_window, &clientRect);

    HDC hdcMem = nullptr;
    HPAINTBUFFER bufferedPaint = BeginBufferedPaint(hdc, &clientRect, BPBF_TOPDOWNDIB, nullptr, &hdcMem);
    if (bufferedPaint)
    {
        DrawBackdrop(hdcMem, clientRect);
        DrawGridLines(hdcMem, clientRect);
        DrawActiveZoneSet(hdcMem, clientRect);
        DrawZoneBuilder(hdcMem, clientRect);
        DrawSwitchButtons(hdcMem, clientRect);
        EndBufferedPaint(bufferedPaint, TRUE);
    }
}

void ZoneWindow::UpdateGrid(int stepColumns, int stepRows) noexcept
{
    bool const shift = GetAsyncKeyState(VK_SHIFT) & 0x8000;
    bool const control = GetAsyncKeyState(VK_CONTROL) & 0x8000;

    RECT clientRect;
    ::GetClientRect(m_window, &clientRect);
    InflateRect(&clientRect, -m_gridMargins.cx, -m_gridMargins.cy);

    Rect gridRect(clientRect);
    if (control || (!stepColumns && !stepRows))
    {
        // Reset
        m_gridColumns = (gridRect.width / 50) + 1;
        m_gridRows = (gridRect.height / 50) + 1;
    }
    else
    {
        stepColumns = stepColumns * (shift ? 5 : 1);
        stepRows = stepRows * (shift ? 5 : 1);

        m_gridColumns = max(1, m_gridColumns + stepColumns);
        m_gridRows = max(1, m_gridRows + stepRows);
    }

    m_gridWidth = gridRect.width / m_gridColumns;
    m_gridHeight = gridRect.height / m_gridRows;
}

void ZoneWindow::UpdateGridMargins(int inc) noexcept
{
    bool const shift = GetAsyncKeyState(VK_SHIFT) & 0x8000;
    bool const control = GetAsyncKeyState(VK_CONTROL) & 0x8000;
    if (control)
    {
        m_gridMargins.cx = 0;
        m_gridMargins.cy = 0;
    }
    else
    {
        inc = inc * (shift ? 5 : 1);
        m_gridMargins.cx = max(0, m_gridMargins.cx + inc);
        m_gridMargins.cy = max(0, m_gridMargins.cy + inc);
    }
    UpdateGrid(0, 0);

    RegistryHelpers::SetValue<SIZE>(m_uniqueId, L"GridMargins", m_gridMargins, sizeof(m_gridMargins));
}

void ZoneWindow::EnterEditorMode() noexcept
{
    MakeActiveZoneSetCustom();
    m_editorMode = true;
}

void ZoneWindow::ExitEditorMode() noexcept
{
    m_editorMode = false;
    if (m_activeZoneSet)
    {
        m_activeZoneSet->Save();
    }
}

void ZoneWindow::OnKeyUp(_In_ WPARAM wparam) noexcept
{
    bool fRedraw = false;

    if ((wparam >= '0') && (wparam<= '9'))
    {
        CycleActiveZoneSetInternal(wparam);
    }
    else
    {
        switch (wparam)
        {
            case VK_DELETE:
            case 'd':
            case 'D':
            {
                // Delete active zone set
                for (auto iter = m_zoneSets.begin(); iter != m_zoneSets.end(); iter++)
                {
                    if (iter->Get() == m_activeZoneSet.Get())
                    {
                        RegistryHelpers::DeleteZoneSet(m_workArea, m_activeZoneSet->GetId());
                        m_zoneSets.erase(iter);
                        m_activeZoneSet = nullptr;
                        break;
                    }
                }
            }
            break;

            case 'r':
            case 'R':
            {
                // Reset zone sets for current work area
                m_zoneSets.clear();
                m_activeZoneSet = nullptr;
                RegistryHelpers::DeleteAllZoneSets(m_workArea);
                InitializeZoneSets();
            }
            break;

            case 'e':
            case 'E':
            {
                // Toggle editor mode
                m_editorMode ? ExitEditorMode() : EnterEditorMode();
            }
            break;

            case 'c':
            case 'C':
            {
                // Create a custom zone
                auto zoneSet = AddZoneSet(ZoneSetLayout::Custom, 0, 0, 0);
                if (zoneSet)
                {
                    UpdateActiveZoneSet(zoneSet.Get());
                }
            }
            break;

            case VK_LEFT: UpdateGrid(-1, 0); break;
            case VK_RIGHT: UpdateGrid(1, 0); break;

            case VK_UP: UpdateGrid(0, 1); break;
            case VK_DOWN: UpdateGrid(0, -1); break;

            case VK_PRIOR: UpdateGridMargins(10); break;
            case VK_NEXT: UpdateGridMargins(-10); break;

            case VK_ESCAPE: g_app->ToggleZoneViewers(); break;
        }
    }
    InvalidateRect(m_window, nullptr, true);
}

ComPtr<IZone> ZoneWindow::ZoneFromPoint(POINT pt) noexcept
{
    if (m_activeZoneSet)
    {
        return m_activeZoneSet->ZoneFromPoint(pt);
    }
    return nullptr;
}

void ZoneWindow::ChooseDefaultActiveZoneSet() noexcept
{
    MONITORINFO mi{};
    mi.cbSize = sizeof(mi);
    if (GetMonitorInfoW(m_monitor, &mi))
    {
        Rect const monitorRect(mi.rcMonitor);

        if ((monitorRect.width == 3840) && (monitorRect.height == 2160))
        {
            ComPtr<IZoneSet> zoneSetBest;
            for (auto zoneSet : m_zoneSets)
            {
                auto zones = zoneSet->GetZones();
                if (zones.size() == 5)
                {
                    if (!zoneSetBest)
                    {
                        zoneSetBest = zoneSet;
                    }
                    else if (zoneSet->GetLayout() == ZoneSetLayout::Focus)
                    {
                        zoneSetBest = zoneSet;
                        break;
                    }
                }
            }

            if (zoneSetBest)
            {
                UpdateActiveZoneSet(zoneSetBest.Get());
            }
        }
        else if (monitorRect.aspectRatio < 40) // ultrawide
        {
            ComPtr<IZoneSet> zoneSetBest;
            for (auto zoneSet : m_zoneSets)
            {
                auto zones = zoneSet->GetZones();
                if (zones.size() == 3)
                {
                    if (!zoneSetBest)
                    {
                        zoneSetBest = zoneSet;
                    }
                    else if (zoneSet->GetLayout() == ZoneSetLayout::Row)
                    {
                        zoneSetBest = zoneSet;
                        break;
                    }
                }
            }

            if (zoneSetBest)
            {
                UpdateActiveZoneSet(zoneSetBest.Get());
            }
        }
        else
        {
            for (auto zoneSet : m_zoneSets)
            {
                auto zones = zoneSet->GetZones();
                if (zones.size() == 1)
                {
                    UpdateActiveZoneSet(zoneSet.Get());
                    break;
                }
            }
        }
    }
}

bool ZoneWindow::IsOccluded(POINT pt, size_t index) noexcept
{
    auto zones = m_activeZoneSet->GetZones();
    size_t i = 1;

    for (auto iter = zones.begin(); iter != zones.end(); iter++)
    {
        ComPtr<IZone> zone;
        if (SUCCEEDED(iter->As(&zone)))
        {
            if (i < index)
            {
                if (PtInRect(&zone->GetZoneRect(), pt))
                {
                    return true;
                }
            }
        }
        i++;
    }
    return false;
}

void ZoneWindow::CycleActiveZoneSetInternal(DWORD wparam) noexcept
{
    if (!m_editorMode)
    {
        if (m_keyLast != wparam)
        {
            m_keyCycle = 0;
        }

        m_keyLast = wparam;

        bool loopAround = true;
        size_t const val = static_cast<size_t>(wparam - L'0');
        size_t i = 0;
        for (auto zoneSet : m_zoneSets)
        {
            if (zoneSet->GetZones().size() == val)
            {
                if (i < m_keyCycle)
                {
                    i++;
                }
                else
                {
                    UpdateActiveZoneSet(zoneSet.Get());
                    loopAround = false;
                    break;
                }
            }
        }

        if ((m_keyCycle > 0) && loopAround)
        {
            // Cycling through a non-empty group and hit the end
            m_keyCycle = 0;
            OnKeyUp(wparam);
        }
        else
        {
            m_keyCycle++;
        }

        g_app->MoveWindowsOnActiveZoneSetChange();
        m_highlightZone = nullptr;
    }
}

void ZoneWindow::FlashZones(bool delayed) noexcept
{
    m_flashMode = true;
    m_flashTimer = SetTimer(m_window, 1, delayed ? 1000 : 1, nullptr);
}

int ZoneWindow::GetSwitchButtonIndexFromPoint(POINT ptClient) noexcept
{
    auto const switchButtonIndex = ((ptClient.x - m_switchButtonContainerRect.left) / (m_switchButtonWidth + m_switchButtonPadding)) + 1;
    return ((switchButtonIndex > 0) && (switchButtonIndex < 10)) ? switchButtonIndex : -1;
}

typedef BOOL(WINAPI *GetDpiForMonitorInternalFunc)(HMONITOR, UINT, UINT*, UINT*);
UINT ZoneWindow::GetDpiForMonitor() noexcept
{
    UINT dpi{};
    if (auto user32 = LoadLibrary(L"user32.dll"))
    {
        if (auto func = reinterpret_cast<GetDpiForMonitorInternalFunc>(GetProcAddress(user32, "GetDpiForMonitorInternal")))
        {
            func(m_monitor, 0, &dpi, &dpi);
        }
        FreeLibrary(user32);
    }

    if (dpi == 0)
    {
        if (HDC hdc = GetDC(nullptr))
        {
            dpi = GetDeviceCaps(hdc, LOGPIXELSX);
            ReleaseDC(nullptr, hdc);
        }
    }

    return (dpi == 0) ? 96 : dpi;
}
#pragma endregion

#pragma region Very interesting stuff
LRESULT CALLBACK ZoneWindow::s_WndProc(_In_ HWND window, UINT message, _In_ WPARAM wparam, _In_ LPARAM lparam) noexcept
{
    auto thisRef = reinterpret_cast<ZoneWindow*>(GetWindowLongPtr(window, GWLP_USERDATA));
    if ((thisRef == nullptr) && (message == WM_CREATE))
    {
        auto createStruct = reinterpret_cast<LPCREATESTRUCT>(lparam);
        thisRef = reinterpret_cast<ZoneWindow*>(createStruct->lpCreateParams);
        thisRef->m_window = window;
        SetWindowLongPtr(window, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(thisRef));
    }

    return (thisRef != nullptr) ? thisRef->WndProc(message, wparam, lparam) :
        DefWindowProc(window, message, wparam, lparam);
}

ComPtr<IZoneWindow> MakeZoneWindow(_In_ HMONITOR monitor, _In_ PCWSTR deviceId) noexcept
{
    return Make<ZoneWindow>(monitor, deviceId);
}
#pragma endregion