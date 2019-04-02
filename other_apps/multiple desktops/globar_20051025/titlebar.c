//
//
//
#define WINVER 0x0501
#define _WIN32_WINNT 0x0501
#include <windows.h>
#include <tchar.h>

#include <uxtheme.h>
#include <tmschema.h>

#include "titlebar.h"

BOOL	 FreeThemeBitmaps();
BOOL	 LoadThemeBitmaps();
HTHEME  _OpenThemeData(HWND, LPCWSTR);
HRESULT _CloseThemeData(HTHEME);


static TCHAR szPropName[] = _T("NewStylieCustomCaptionSubclassPtr");

#define B_EDGE				2


extern HWND	g_hwndNotify;

// defined in themes.c
extern HBITMAP  g_hThemeBitmap;
extern HDC		g_hThemeBitmapDC;
extern COLORREF g_GlyphColor;


typedef struct
{
	UINT	uCmd;			//command to send when clicked
	int		nRightBorder;	//Pixels between this button and buttons to the right
	HBITMAP hBmp;			//Bitmap to display

//	BOOL	fPressed;		//Private.
//	BOOL	fMouseOver;		//Private.
	
} TITLEBUT;

typedef struct
{
	TITLEBUT	buttons[MAX_TITLE_BUTTONS];
	int			nNumButtons;
	WNDPROC		wpOldProc;

	BOOL		fMouseDown;
	BOOL		fMouseOver;

	int			iActiveButton;
	BOOL		fCaptionActive;

} CUSTCAPTION;

static CUSTCAPTION *GetCustomCaption(HWND hwnd)
{
	return (CUSTCAPTION *)GetProp(hwnd, szPropName);
}

static void RedrawNC(HWND hwnd)
{
	SetWindowPos(hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_DRAWFRAME);
}

static int CalcTopEdge(HWND hwnd)
{
	DWORD dwStyle = GetWindowLong(hwnd, GWL_STYLE);

	if(dwStyle & WS_THICKFRAME)
		return GetSystemMetrics(SM_CYSIZEFRAME);
	else
		return GetSystemMetrics(SM_CYFIXEDFRAME);
}

static int CalcRightEdge(HWND hwnd)
{
	DWORD dwStyle = GetWindowLong(hwnd, GWL_STYLE);

	if(dwStyle & WS_THICKFRAME)
		return GetSystemMetrics(SM_CXSIZEFRAME);
	else
		return GetSystemMetrics(SM_CXFIXEDFRAME);

}

WNDPROC SafeSubclassWindow(HWND hwnd, WNDPROC NewProc)
{
	if(IsWindowUnicode(hwnd))
		return (WNDPROC)SetWindowLongPtrW(hwnd, GWL_WNDPROC, (LONG_PTR)NewProc);
	else
		return (WNDPROC)SetWindowLongPtrA(hwnd, GWL_WNDPROC, (LONG_PTR)NewProc);
}

LRESULT SafeCallWndProc(WNDPROC OldProc, HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	if(IsWindowUnicode(hwnd))
		return CallWindowProcW(OldProc, hwnd, msg, wParam, lParam);
	else
		return CallWindowProcA(OldProc, hwnd, msg, wParam, lParam);
}


//
//	Return the starting position of the first inserted button.
//	We need to take into account the size of the close button,
//	The minimize, maximize buttons etc
//
//	-------------------------------------------------\
//                                        [-][+] [X] |
//                                       ^
//                                       |
//                   Return is pos       |--- Ret  --|
//
//
static int GetRightEdgeOffset(CUSTCAPTION *ctp, HWND hwnd)
{
	DWORD dwStyle   = GetWindowLong(hwnd, GWL_STYLE);
	DWORD dwExStyle = GetWindowLong(hwnd, GWL_EXSTYLE);

	int nButSize = 0;
	int nSysButSize;

	if(dwExStyle & WS_EX_TOOLWINDOW)
	{
		nSysButSize = GetSystemMetrics(SM_CXSMSIZE) - B_EDGE;

		if(dwStyle & WS_SYSMENU)	
			nButSize += nSysButSize + B_EDGE;

		return nButSize + CalcRightEdge(hwnd);
	}
	else
	{
		nSysButSize = GetSystemMetrics(SM_CXSIZE) - B_EDGE;

		// Window has Close [X] button. This button has a 2-pixel
		// border on either size
		if(dwStyle & WS_SYSMENU)	
		{
			nButSize += nSysButSize + B_EDGE;
		}

		// If either of the minimize or maximize buttons are shown,
		// Then both will appear (but may be disabled)
		// This button pair has a 2 pixel border on the left
		if(dwStyle & (WS_MINIMIZEBOX | WS_MAXIMIZEBOX) )
		{
			nButSize += B_EDGE + nSysButSize * 2;
		}
		// A window can have a question-mark button, but only
		// if it doesn't have any min/max buttons
		else if(dwExStyle & WS_EX_CONTEXTHELP)
		{
			nButSize += B_EDGE + nSysButSize;
		}
		
		// Now calculate the size of the border...aggghh!
		return nButSize + CalcRightEdge(hwnd);
	}
}


//
//	Input  - rect is coords of window rectangle
//			 (in screen or relative coords)
//
//	Output - rect is adjusted to describe button rectangle
//
static void GetButtonRect(CUSTCAPTION *ctp, HWND hwnd, int idx, RECT *rect, BOOL fWindowRelative)
{
	int i, re_start;
	int cxBut, cyBut;

	if(GetWindowLong(hwnd, GWL_EXSTYLE) & WS_EX_TOOLWINDOW)
	{
		cxBut = GetSystemMetrics(SM_CXSMSIZE);
		cyBut = GetSystemMetrics(SM_CYSMSIZE);
	}
	else
	{
		cxBut = GetSystemMetrics(SM_CXSIZE);
		cyBut = GetSystemMetrics(SM_CYSIZE);
	}

	// right-edge starting point of inserted buttons
	re_start = GetRightEdgeOffset(ctp, hwnd);

	GetWindowRect(hwnd, rect);
	
	if(fWindowRelative)
		OffsetRect(rect, -rect->left, -rect->top);

	//Find the correct button - but take into
	//account all other buttons.
	for(i = 0; i <= idx; i++)
	{
		re_start += ctp->buttons[i].nRightBorder + cxBut - B_EDGE;
	}

	rect->left   = rect->right  - re_start;
	rect->top    = rect->top  + CalcTopEdge(hwnd) +  B_EDGE;
	rect->right  = rect->left + cxBut - B_EDGE;
	rect->bottom = rect->top  + cyBut - B_EDGE*2;
}

static LRESULT Caption_NCHitTest(CUSTCAPTION *ctp, HWND hwnd, WPARAM wParam, LPARAM lParam)
{
	RECT  rect;
	POINT pt;
	int   i;
	UINT  ret;

	pt.x = (short)LOWORD(lParam);
	pt.y = (short)HIWORD(lParam);

	ret = SafeCallWndProc(ctp->wpOldProc, hwnd, WM_NCHITTEST, wParam, lParam);

	//If the mouse is in the caption, then check to
	//see if it is over one of our buttons
	if(ret == HTCAPTION)
	{
		for(i = 0; i < ctp->nNumButtons; i++)
		{
			GetButtonRect(ctp, hwnd, i, &rect, FALSE);
			InflateRect(&rect, 0, B_EDGE);

			//If the mouse is in any custom button, then
			//We need to override the default behaviour.
			if(PtInRect(&rect, pt))
			{
				return HTBORDER;
			}
		}
	}

	return ret;
}

//
//	If hBitmap is MONOCHROME then 
//	whites will be drawn transparently, 
//  blacks will be drawn normally
//	So, it will look just like a caption button
//
//	If hBitmap is > 2 colours, then no transparent
//  drawing will take place....i.e. DIY!
//
static void CenterBitmap(HTHEME hTheme, HDC hdc, RECT *rc, HBITMAP hBitmap)
{
	BITMAP bm;
	int cx;
	int cy;   
	HDC memdc;
	HBITMAP hOldBM;
	RECT  rcDest = *rc;   
	POINT p;
	SIZE  delta;
	COLORREF colorOld, colorNew;

	if(hBitmap == NULL) 
		return;

	// center bitmap in caller's rectangle   
	GetObject(hBitmap, sizeof bm, &bm);   
	
	cx = bm.bmWidth;
	cy = bm.bmHeight;

	delta.cx = (rc->right-rc->left - cx) / 2;
	delta.cy = (rc->bottom-rc->top - cy) / 2;
	
	if(rc->right-rc->left > cx)
	{
		SetRect(&rcDest, rc->left+delta.cx, rc->top + delta.cy, 0, 0);   
		rcDest.right = rcDest.left + cx;
		rcDest.bottom = rcDest.top + cy;
		p.x = 0;
		p.y = 0;
	}
	else
	{
		p.x = -delta.cx;   
		p.y = -delta.cy;
	}
   
	// select checkmark into memory DC
	memdc = CreateCompatibleDC(hdc);
	hOldBM = (HBITMAP)SelectObject(memdc, hBitmap);
   
	
	// set BG color based on selected state   
	if(hTheme)
	{
		//GetThemeColor(hTheme, WP_MAXBUTTON, MAXBS_NORMAL, TMT_BORDERCOLORHINT, &colorNew);
		//GetThemeColor(hTheme, WP_SYSBUTTON, SBS_NORMAL, TMT_BORDERCOLORHINT, &colorNew);
		colorNew = g_GlyphColor;//GetPixel(hdc, rcDest.left, rcDest.top);
	}	
	else
		colorNew = GetSysColor(COLOR_3DDKSHADOW);

	colorOld = SetTextColor(hdc, colorNew);

	//BitBlt
	MaskBlt
		(hdc, rcDest.left, 
				rcDest.top, 
				rcDest.right-rcDest.left, 
				rcDest.bottom-rcDest.top, 
				memdc, 
				p.x, 
				p.y, 
				hBitmap, 
				p.x, 
				p.y, 
				//SRCCOPY
				MAKEROP4(
				0x00AA0029, SRCCOPY)
				);

	// restore
	SetTextColor(hdc, colorOld);
	SelectObject(memdc, hOldBM);
	DeleteDC(memdc);

}

void DrawButtonBackground(HTHEME hTheme, HDC hdc, RECT *rect, BOOL fPressed, BOOL fMouseOver, BOOL fActive)
{
	if(hTheme)
	{
		int		nThemeBitmapIdx = 0;
		BITMAP bmp;
		RECT rrr;

		LoadThemeBitmaps(WindowFromDC(hdc));


		GetThemeRect(hTheme, WP_MAXBUTTON, MAXBS_NORMAL, TMT_DEFAULTPANESIZE, &rrr);

		if(fMouseOver && !fPressed)
			nThemeBitmapIdx = 1;

		if(fPressed)
			nThemeBitmapIdx = 2;

		if(fActive == FALSE)
			nThemeBitmapIdx += 4;

		GetObject(g_hThemeBitmap, sizeof(BITMAP), &bmp);
	//	GetThemePartSize(hTheme, hdc, WP_MAXBUTTON, MAXBS_PUSHED, NULL, TS_TRUE, &sz);
	//	GetThemeColor(hTheme, WP_MAXBUTTON, MAXBS_PUSHED, TMT_BORDERCOLORHINT, &col);


		//DrawThemeParentBackground(WindowFromDC(hdc), hdc, rect);

	//	DrawThemeBackground(hTheme, hdc,
	//	WP_MAXBUTTON, MAXBS_PUSHED, rect, rect);
		//FillRect(hdc, rect, GetStockObject(BLACK_BRUSH));

		TransparentBlt(
		//StretchBlt(
			hdc, 
			rect->left, 
			rect->top, 
			rect->right-rect->left,
			rect->bottom-rect->top,
			g_hThemeBitmapDC,
			0,
			(bmp.bmHeight / 8) * nThemeBitmapIdx,
			bmp.bmWidth,
			bmp.bmHeight / 8,
			//RGB(255,0,255)
			GetPixel(g_hThemeBitmapDC, 0, (bmp.bmHeight / 8) * nThemeBitmapIdx)
			//SRCCOPY
			);	
	}
	else
	{
		UINT uButType = DFCS_BUTTONPUSH | (fPressed ? DFCS_PUSHED : 0);
		DrawFrameControl(hdc, rect, DFC_BUTTON,  uButType);
	}
}

static LRESULT Caption_NCPaint(CUSTCAPTION *ctp, HWND hwnd, HRGN hrgn)
{
	RECT	rect, rect1;
	BOOL	fRegionOwner = FALSE;
	int		i;
	HDC		hdc;
	int		x, y;

	HRGN	hrgn1;

	HTHEME	hTheme = _OpenThemeData(hwnd, L"window");

	GetWindowRect(hwnd, &rect);

	x = rect.left;
	y = rect.top;

	//Create a region which covers the whole window. This
	//must be in screen coordinates
	if(hrgn == (HRGN)1 || hrgn == 0)
	{
		hrgn = CreateRectRgnIndirect(&rect);
		fRegionOwner = TRUE;
	}

	// Clip our custom buttons out of the way...
	for(i = 0; i < ctp->nNumButtons; i++)
	{
		//Get button rectangle in screen coords
		GetButtonRect(ctp, hwnd, i, &rect1, FALSE);

		hrgn1 = CreateRectRgnIndirect(&rect1);

		//Cut out a button-shaped hole
		CombineRgn(hrgn, hrgn, hrgn1, RGN_XOR);

		DeleteObject(hrgn1);
	}

	//
	//	Call the default window procedure with our modified window region!
	//	(REGION MUST BE IN SCREEN COORDINATES)
	//
	SafeCallWndProc(ctp->wpOldProc, hwnd, WM_NCPAINT, (WPARAM)hrgn, 0);
	
	//
	//	Now paint our custom buttons in the holes that are
	//	left by our clipping. All drawing in the Non-client area
	//  is window-relative (Not in screen coords)
	//

	hdc = GetWindowDC(hwnd);

	// Draw buttons in a loop
	for(i = 0; i < ctp->nNumButtons; i++)
	{
		BOOL fMouseDown;
		BOOL fMouseOver;

		//Get Button rect in window coords
		GetButtonRect(ctp, hwnd, i, &rect1, TRUE);
		
		//GetThemeRect(hTheme, WP_MAXBUTTON, MAXBS_PUSHED, TMT_DEFAULTPANESIZE, &rect1);
	
		if(i == ctp->iActiveButton)
		{
			fMouseDown = ctp->fMouseDown && ctp->fMouseOver;
			fMouseOver = ctp->fMouseOver;
		}
		else
		{
			fMouseDown = FALSE;
			fMouseOver = FALSE;
		}

		//
		//	Draw button background!!
		// 
		DrawButtonBackground(hTheme, hdc, &rect1, 
			fMouseDown, 
			fMouseOver,
			ctp->fCaptionActive
			);
			
		InflateRect(&rect1, -2, -2);
		rect1.right--;
		rect1.bottom--;

		// only offset the glyph if we're not themed
		if(fMouseDown && hTheme == 0)
			OffsetRect(&rect1, 1, 1);
			
		//
		//	Draw button foreground!!
		//
		CenterBitmap(hTheme, hdc, &rect1, ctp->buttons[i].hBmp);
	}

	ReleaseDC(hwnd, hdc);
	
	if(fRegionOwner)
		DeleteObject(hrgn);

	_CloseThemeData(hTheme);
	
	return 0;
}

/*
static LRESULT Caption_NCPaint2(CustomCaption *ctp, HWND hwnd, HRGN hrgn)
{
	HTHEME hTheme = 0;//OpenThemeData(hwnd, L"window");
	HDC hdc = GetWindowDC(hwnd);

	RECT rect, clip;
	RECT rect1;
	GetWindowRect(hwnd, &rect);
	
	OffsetRect(&rect, -rect.left, -rect.top);
	CopyRect(&clip, &rect);
	
	rect1.left = 0;
	rect1.top = 0;
	rect1.right=100;
	rect1.bottom=30;
	
	if(0)
	{
		DrawFrameControl(hdc, &rect1, DFC_BUTTON,  DFCS_PUSHED);
	}
	else
	{
		//DrawThemeBackground(hTheme, hdc, WP_MAXBUTTON, CS_ACTIVE, &rect1, &rect1);

	}

	ReleaseDC(hwnd, hdc);

	CloseThemeData(hTheme);
	
	//DefWindowProc(hwnd, WM_NCPAINT, hrgn, 0);
	SafeCallWndProc(ctp->wpOldProc, hwnd, WM_NCPAINT, (WPARAM)hrgn, 0);


	DrawThemeBackground(hTheme, hdc, WP_MAXBUTTON, CS_ACTIVE, &rect1, &rect1);
}
*/

//
//	This is a generic message handler used by WM_SETTEXT and WM_NCACTIVATE.
//	It works by turning off the WS_VISIBLE style, calling
//	the original window procedure, then turning WS_VISIBLE back on.
//
//	This prevents the original wndproc from redrawing the caption.
//	Last of all, we paint the caption ourselves with the inserted buttons
//
static LRESULT Caption_Wrapper(CUSTCAPTION *ctp, HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	UINT ret;
	DWORD dwStyle;

	dwStyle = GetWindowLong(hwnd, GWL_STYLE);

	//Turn OFF WS_VISIBLE, so that WM_NCACTIVATE does not
	//paint our window caption...
	SetWindowLong(hwnd, GWL_STYLE, dwStyle & ~WS_VISIBLE);
	
	//Do the default thing..
	ret = SafeCallWndProc(ctp->wpOldProc, hwnd, msg, wParam, lParam);
	
	//Restore the original style
	SetWindowLong(hwnd, GWL_STYLE, dwStyle);

	return ret;
}

static LRESULT Caption_NCActivate(CUSTCAPTION *ctp, HWND hwnd, WPARAM wParam, LPARAM lParam)
{
	LRESULT lResult = Caption_Wrapper(ctp, hwnd, WM_NCACTIVATE, wParam, lParam);

	ctp->fCaptionActive = wParam;

	//Paint the whole window frame + caption
	Caption_NCPaint(ctp, hwnd, (HRGN)1);

	return lResult;	
}

static LRESULT Caption_SetText(CUSTCAPTION *ctp, HWND hwnd, WPARAM wParam, LPARAM lParam)
{
	LRESULT lResult = Caption_Wrapper(ctp, hwnd, WM_SETTEXT, wParam, lParam);

	//Paint the whole window frame + caption
	Caption_NCPaint(ctp, hwnd, (HRGN)1);

	return lResult;
}

//
//	NonClient Left Button Down.
//	Mouse is in screen coordinates
//
static LRESULT Caption_NCLButtonDown(CUSTCAPTION *ctp, HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	int i;
	RECT rect;
	POINT pt;
	
	pt.x = (short)LOWORD(lParam);
	pt.y = (short)HIWORD(lParam);

	//If mouse has been clicked in caption
	//(Note: the NCHITTEST handler changes HTCAPTION to HTBORDER
	//       if we are over an inserted button)
	if(wParam == HTBORDER)
	{
		for(i = 0; i < ctp->nNumButtons; i++)
		{
			GetButtonRect(ctp, hwnd, i, &rect, FALSE);
			InflateRect(&rect, 0, 2);
			
			//if clicked in a custom button
			if(PtInRect(&rect, pt))
			{
				ctp->iActiveButton = i;
				ctp->fMouseDown = TRUE;
				ctp->fMouseOver = TRUE;
				
				SetCapture(hwnd);
				
				RedrawNC(hwnd);
				
				return 0;
			}
		}
	}

	return SafeCallWndProc(ctp->wpOldProc, hwnd, msg, wParam, lParam);
}

//
//	Left-button UP. Coords are CLIENT relative
//
static LRESULT Caption_LButtonUp(CUSTCAPTION *ctp, HWND hwnd, WPARAM wParam, LPARAM lParam)
{
	RECT rect;
	POINT pt;
	
	pt.x = (short)LOWORD(lParam);
	pt.y = (short)HIWORD(lParam);
	ClientToScreen(hwnd, &pt);

	if(ctp->fMouseDown)
	{
		ReleaseCapture();

		GetButtonRect(ctp, hwnd, ctp->iActiveButton, &rect, FALSE);
		InflateRect(&rect, 0, 2);

		//if clicked in a custom button
		if(PtInRect(&rect, pt))
		{
			UINT uCmd = ctp->buttons[ctp->iActiveButton].uCmd;
			PostMessage(g_hwndNotify, WM_COMMAND, MAKEWPARAM(uCmd, 0), (LPARAM)hwnd);
		}

		//ctp->buttons[ctp->iActiveButton].fPressed = FALSE;
		ctp->fMouseDown = FALSE;

		RedrawNC(hwnd);
						
		return 0;
	}

	return SafeCallWndProc(ctp->wpOldProc, hwnd, WM_LBUTTONUP, wParam, lParam);
}

static LRESULT Caption_MouseMove(CUSTCAPTION *ctp, HWND hwnd, WPARAM wParam, LPARAM lParam)
{
	RECT rect;
	POINT pt;
	BOOL fPressed;
	
	pt.x = (short)LOWORD(lParam);
	pt.y = (short)HIWORD(lParam);
	ClientToScreen(hwnd, &pt);
	
	if(ctp->fMouseDown)
	{
		GetButtonRect(ctp, hwnd, ctp->iActiveButton, &rect, FALSE);
		InflateRect(&rect, 0, 2);
		
		fPressed = PtInRect(&rect, pt);
		
		if(fPressed != ctp->fMouseOver)
		{
			ctp->fMouseOver = fPressed;
			RedrawNC(hwnd);
		}
		
		return 0;
	}

	return SafeCallWndProc(ctp->wpOldProc, hwnd, WM_MOUSEMOVE, wParam, lParam);
}

static LRESULT Caption_NcMouseMove(CUSTCAPTION *ctp, HWND hwnd, WPARAM wParam, LPARAM lParam)
{
	RECT rect;
	POINT pt;
	int  i;
	int  iActiveButton = -1;
	BOOL fMouseOver    = FALSE;
	
	pt.x = (short)LOWORD(lParam);
	pt.y = (short)HIWORD(lParam);
	
	if(wParam == HTBORDER)
	{
		for(i = 0; i < ctp->nNumButtons; i++)
		{
			GetButtonRect(ctp, hwnd, i, &rect, FALSE);
			InflateRect(&rect, 0, 2);
			
			//if moved over a custom button			
			if(PtInRect(&rect, pt))
			{
				fMouseOver = TRUE;
				iActiveButton = i;		
			}
		}
	}

	//ctp->buttons[ctp->iActiveButton].fMouseOver	= FALSE;

	if((iActiveButton != ctp->iActiveButton && iActiveButton != -1)
		|| fMouseOver != ctp->fMouseOver)
	{
		ctp->fMouseOver = fMouseOver;

		if(iActiveButton != -1)
			ctp->iActiveButton = iActiveButton;
				
		RedrawNC(hwnd);			
	}

	return SafeCallWndProc(ctp->wpOldProc, hwnd, WM_NCMOUSEMOVE, wParam, lParam);
}


//Replacement window procedure
static LRESULT CALLBACK NewWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	CUSTCAPTION *ctp = GetCustomCaption(hwnd);
	WNDPROC oldproc = ctp->wpOldProc;
	
	switch(msg)
	{
	//clean up when window is destroyed
	case WM_NCDESTROY:
		HeapFree(GetProcessHeap(), 0, ctp);
		break;

	case WM_NCHITTEST:
		return Caption_NCHitTest(ctp, hwnd, wParam, lParam);

	//These three messages all cause the caption to
	//be repainted, so we have to handle all three to properly
	//support inserted buttons
	case WM_NCACTIVATE:
		return Caption_NCActivate(ctp, hwnd, wParam, lParam);

	case WM_SETTEXT:
		return Caption_SetText(ctp, hwnd, wParam, lParam);

	case WM_NCPAINT:
		return Caption_NCPaint(ctp, hwnd, (HRGN)wParam);

	//Mouse support
	case WM_NCLBUTTONDBLCLK:
	case WM_NCLBUTTONDOWN:
		return Caption_NCLButtonDown(ctp, hwnd, msg, wParam, lParam);

	case WM_LBUTTONUP:
		return Caption_LButtonUp(ctp, hwnd, wParam, lParam);

	case WM_MOUSEMOVE:
		return Caption_MouseMove(ctp, hwnd, wParam, lParam);

	// added for theme support
	case WM_NCMOUSEMOVE:
		return Caption_NcMouseMove(ctp, hwnd, wParam, lParam);
		
	case WM_NCMOUSELEAVE:
		Caption_NcMouseMove(ctp, hwnd, HTBORDER, GetMessagePos());
		Caption_NCPaint(ctp, hwnd, (HRGN)1);
		break;

	case WM_THEMECHANGED:
		FreeThemeBitmaps();
		LoadThemeBitmaps();
		break;

	}

	//call the old window procedure
	return SafeCallWndProc(oldproc, hwnd, msg, wParam, lParam);
}


//
//	Insert a button into specified window's titlebar
//
BOOL WINAPI Caption_InsertButton(HWND hwnd, UINT uCmd, int nBorder, HBITMAP hBmp)
{
	CUSTCAPTION *ctp = GetCustomCaption(hwnd);
	int idx;

	// for XP-theme support, load the theme titlebar bitmaps
	LoadThemeBitmaps(hwnd);

	// If this window doesn't have any buttons yet,
	// then perform the subclass and allocate structures etc. 
	if(ctp == 0)
	{
		TRACKMOUSEEVENT trackmouse = { sizeof(TRACKMOUSEEVENT), TME_NONCLIENT, hwnd };

		//allocate memory for our subclass information
		ctp = HeapAlloc(GetProcessHeap(), 0, sizeof(CUSTCAPTION));

		ctp->nNumButtons    = 0;
		ctp->fMouseDown     = FALSE;
		ctp->iActiveButton  = 0;
		ctp->fMouseOver		= FALSE;
		ctp->fCaptionActive = GetActiveWindow() == hwnd ? TRUE : FALSE;

		//assign this to the window in question
		SetProp(hwnd, szPropName, (HANDLE)ctp);

		//subclass the window
		ctp->wpOldProc = SafeSubclassWindow(hwnd, NewWndProc);

		TrackMouseEvent(&trackmouse);
	}

	idx = ctp->nNumButtons++;

	ctp->buttons[idx].hBmp			= hBmp;
	ctp->buttons[idx].nRightBorder  = nBorder;
	ctp->buttons[idx].uCmd			= uCmd;
	
	return TRUE;
}

//
//	Remove the last inserted button from window's titlebar
//
BOOL WINAPI Caption_RemoveAllButtons(HWND hwnd)
{
	CUSTCAPTION *ctp = GetCustomCaption(hwnd);

	if(ctp == 0)
		return FALSE;

	// remove the subclass
	RemoveProp(hwnd, szPropName);
	
	SafeSubclassWindow(hwnd, ctp->wpOldProc);

	HeapFree(GetProcessHeap(), 0, ctp);

	FreeThemeBitmaps();

	RedrawNC(hwnd);

	return TRUE;
}