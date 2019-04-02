//
//	themes.c	
//
//	xp-theme support for custom titlebars
//
//	The problem:
//
//	We need to draw a custom titlebar button on an xp-theme titlebar.
//	However we don't want to store custom bitmaps for every type of theme
//	that Microsoft/3rd parties might release....i.e. we want to draw our
//  buttons using the current theme, but with modified glyphs
//
//	The solution is to manually load the Theme DLL used by the current
//  theme and extract the blank titlebar button bitmaps.
//
//
#define WINVER 0x0501
#define _WIN32_WINNT 0x0501
#include <windows.h>
#include <tchar.h>

#include <uxtheme.h>
#include <tmschema.h>
#include "titlebar.h"

#pragma comment(lib,    "Uxtheme.lib")
#pragma comment(lib,    "Delayimp.lib")
#pragma comment(linker, "/delayload:Uxtheme.dll")
#pragma comment(lib,	"MSIMG32.lib")

BOOL	 PreloadBitmaps();
HBITMAP  g_hThemeBitmap;
HDC		 g_hThemeBitmapDC;
HANDLE   g_hOldThemeBmp;
HMODULE  g_hThemeModule;
COLORREF g_GlyphColor;

BOOL	g_fThemeApiAvailable = FALSE;

HTHEME _OpenThemeData(HWND hwnd, LPCWSTR pszClassList)
{
	if(g_fThemeApiAvailable)
		return OpenThemeData(hwnd, pszClassList);
	else
		return NULL;
}

HRESULT _CloseThemeData(HTHEME hTheme)
{
	if(g_fThemeApiAvailable)
		return CloseThemeData(hTheme);
	else
		return E_FAIL;
}

VOID InitThemeSupport()
{
	HMODULE hThemeAPI = LoadLibrary("uxtheme.dll");

	if(hThemeAPI &&
		GetProcAddress(hThemeAPI, "OpenThemeData")			&&
		GetProcAddress(hThemeAPI, "CloseThemeData")			&&
		GetProcAddress(hThemeAPI, "GetCurrentThemeName")	&&
		GetProcAddress(hThemeAPI, "GetThemeFilename")
		)
	{
		g_fThemeApiAvailable = TRUE;
	}
	else
	{
		g_fThemeApiAvailable = FALSE;
	}

	FreeLibrary(hThemeAPI);
}

//
//	Convert from a theme bitmap-name to a theme resource-name
//
//	e.g. from:
//
//		Blue\CaptionButton.bmp
//
//	to:
//
//		BLUE_CAPTIONBUTTON_BMP
//
void MakeThemeBitmapResourceName(WCHAR *wszThemeBitmap)
{
	for( ; *wszThemeBitmap; wszThemeBitmap++)
	{
		if(*wszThemeBitmap == '\\')
			*wszThemeBitmap = '_';

		if(*wszThemeBitmap == '.')
			*wszThemeBitmap = '_';
	}
}


BOOL FreeThemeBitmaps()
{
	if(g_hThemeBitmapDC != 0)
	{
		SelectObject(g_hThemeBitmapDC, g_hOldThemeBmp);
		DeleteObject(g_hThemeBitmap);
		DeleteDC(g_hThemeBitmapDC);

		g_hThemeBitmap   = 0;
		g_hThemeBitmapDC = 0;
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}

//
//	
//
BOOL LoadThemeBitmaps(HWND hwnd)
{
	WCHAR	wszThemePath[MAX_PATH];
	WCHAR	wszThemeBitmap[100];
	HDC		hdc;

	HTHEME	hTheme;
	
	// do we already have the bitmaps loaded????
	if(g_hThemeBitmapDC != 0)
		return TRUE;

	// open the theme for this window
	if((hTheme = _OpenThemeData(hwnd, L"window")) == 0)
	{
		return FALSE;
	}

	/// get current theme name
	// (e.g. "c:\windows\resources\Themes\luna\luna.msstyles")
	if(S_OK != GetCurrentThemeName(wszThemePath, MAX_PATH, 0, 0, 0, 0))
	{
		CloseThemeData(hTheme);
		return FALSE;
	}

	// get the name of the bitmap used for caption-button
	if(S_OK != GetThemeFilename(hTheme, WP_MAXBUTTON, MAXBS_NORMAL, TMT_IMAGEFILE, wszThemeBitmap, 100))
	{
		CloseThemeData(hTheme);	
		return FALSE;
	}
	
	// load the theme DLL
	if((g_hThemeModule = LoadLibraryExW(wszThemePath, 0, LOAD_LIBRARY_AS_DATAFILE)) == 0)
	{
		CloseThemeData(hTheme);
		return FALSE;
	}
		
	// convert the theme bitmap name to a resource-name
	// (e.g. "Blue\CaptionButton.bmp" ---> "BLUE_CAPTIONBUTTON_BMP"
	MakeThemeBitmapResourceName(wszThemeBitmap);

	//OutputDebugStringW(wszThemeBitmap);
	//OutputDebugString("   loading new\n");
		
	if((g_hThemeBitmap = LoadImageW(g_hThemeModule, wszThemeBitmap, IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION)) == 0)
	{
		FreeLibrary(g_hThemeModule);
		CloseThemeData(hTheme);
		return FALSE;
	}
		
	//
	//	Create a memory-DC to hold the theme bitmap
	//
	hdc				 = GetDC(0);
	g_hThemeBitmapDC = CreateCompatibleDC(hdc);
	g_hOldThemeBmp   = SelectObject(g_hThemeBitmapDC, g_hThemeBitmap);
	g_GlyphColor	 = GetPixel(g_hThemeBitmapDC, 0, 3);
	ReleaseDC(0, hdc);

	// Free the DLL now the bitmap is loaded!!!
	FreeLibrary(g_hThemeModule);
	CloseThemeData(hTheme);
	
	return FALSE;
}

