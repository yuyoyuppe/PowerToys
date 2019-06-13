#pragma once

#define IsFlagSet(obj, f) (((obj) & (f)) == (f))

struct Rect
{
    Rect() {}

    Rect(RECT rect) :
        m_rect(rect),
        x(rect.left),
        y(rect.top),
        width(rect.right - rect.left),
        height(rect.bottom - rect.top),
        left(rect.left),
        top(rect.top),
        right(rect.right),
        bottom(rect.bottom)
    {
        aspectRatio = height * 100 / width;
    }

    int x{};
    int y{};
    int width{};
    int height{};
    int left{};
    int top{};
    int right{};
    int bottom{};
    int aspectRatio{};

private:
    RECT m_rect{};
};

inline void MakeWindowTransparent(_In_ HWND window)
{
    int const pos = -GetSystemMetrics(SM_CXVIRTUALSCREEN) - 8;
    HRGN hrgn = CreateRectRgn(pos, 0, (pos + 1), 1);
    if (hrgn != nullptr)
    {
        DWM_BLURBEHIND bh = { DWM_BB_ENABLE | DWM_BB_BLURREGION, TRUE, hrgn, FALSE };
        DwmEnableBlurBehindWindow(window, &bh);
        DeleteObject(hrgn);
    }
}

inline void InitRGB(_Out_ RGBQUAD *quad, BYTE alpha, COLORREF color)
{
    ZeroMemory(quad, sizeof(*quad));
    quad->rgbReserved = alpha;
    quad->rgbRed = GetRValue(color) * alpha / 255;
    quad->rgbGreen = GetGValue(color) * alpha / 255;
    quad->rgbBlue = GetBValue(color) * alpha / 255;
}

inline void FillRectARGB(_In_ HDC hdc, _In_ RECT const *prcFill, BYTE alpha, COLORREF color, bool blendAlpha)
{
    BITMAPINFO bi;
    ZeroMemory(&bi, sizeof(bi));
    bi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bi.bmiHeader.biWidth = 1;
    bi.bmiHeader.biHeight = 1;
    bi.bmiHeader.biPlanes = 1;
    bi.bmiHeader.biBitCount = 32;
    bi.bmiHeader.biCompression = BI_RGB;

    RECT fillRect;
    CopyRect(&fillRect, prcFill);
    if ((alpha == 255) || !blendAlpha)
    {
        // Opaque or the caller does not want to blend the alpha
        RGBQUAD bitmapBits;
        InitRGB(&bitmapBits, alpha, color);
        StretchDIBits(
            hdc,
            fillRect.left,
            fillRect.top,
            fillRect.right - fillRect.left,
            fillRect.bottom - fillRect.top,
            0, 0, 1, 1, &bitmapBits, &bi, DIB_RGB_COLORS, SRCCOPY);
    }
    else
    {
        HDC hdcSrc = CreateCompatibleDC(hdc);
        if (hdcSrc != nullptr)
        {
            void *pBitmapBits;
            HBITMAP bitmapSource = CreateDIBSection(hdcSrc, &bi, DIB_RGB_COLORS, &pBitmapBits, nullptr, 0);
            if (bitmapSource != nullptr)
            {
                InitRGB(reinterpret_cast<RGBQUAD *>(pBitmapBits), alpha, color);

                HGDIOBJ bitmapOld = SelectObject(hdcSrc, bitmapSource);
                BLENDFUNCTION bf = { AC_SRC_OVER, 0, 255, AC_SRC_ALPHA };
                GdiAlphaBlend(
                    hdc,
                    fillRect.left,
                    fillRect.top,
                    fillRect.right - fillRect.left,
                    fillRect.bottom - fillRect.top,
                    hdcSrc, 0, 0, 1, 1, bf);
                SelectObject(hdcSrc, bitmapOld);
                DeleteObject(bitmapSource);
            }
            DeleteDC(hdcSrc);
        }
    }
}

inline void FrameRectARGB(HDC hdc, const RECT &rc, BYTE bAlpha, COLORREF clr, int thickness)
{
    RECT sides[] = {
        { rc.left, rc.top, (rc.left + thickness), rc.bottom },
        { (rc.right - thickness), rc.top, rc.right, rc.bottom },
        { (rc.left + thickness), rc.top, (rc.right - thickness), (rc.top + thickness) },
        { (rc.left + thickness), (rc.bottom - thickness), (rc.right - thickness), rc.bottom }
    };

    for (UINT i = 0; i < ARRAYSIZE(sides); i++)
    {
        FillRectARGB(hdc, &(sides[i]), bAlpha, clr, false);
    }
}