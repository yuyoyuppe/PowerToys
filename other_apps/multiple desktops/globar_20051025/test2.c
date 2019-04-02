//
//	globargui.c
//
//	Example C program which uses the globar DLL
// 
//	v1.0 - J Brown
//
//	www.catch22.net
//
//	Uses LoadLibrary to load globar.dll
//	Then does GetProcAddress to get API 
//	
//	BOOL WINAPI AddGlobalButton(HWND hwndNotify, UINT uCommandId, UINT uBitmapResourceId)
//	BOOL WINAPI EnableGlobalButtons(BOOL fEnable)
//
//
//
#include <windows.h>
#include <stdio.h>
#include <stdarg.h>
#include "resource.h"
#include "globar.h"

#ifdef _DEBUG
#pragma comment(lib, "debug\\globar.lib")
#else
#pragma comment(lib, "release\\globar.lib")
#endif

#pragma comment(linker, "/OPT:NOWIN98")

void msgprintf(HWND hwnd, const char *fmt, ...)
{
	va_list vargs;
	char buf[1000];

	va_start(vargs, fmt);

	vsprintf(buf, fmt, vargs);
	MessageBox(hwnd, buf, "Globar Test", MB_OK);

	va_end(vargs);
}

//
//	Silly function to find the capture.dll file - in visual studio,
//	it is always in .\debug or .\release and not in the current directory
//	so this is just a bodge to make building+testing easier
//
/*
void FindModuleFileName(char *szFileName, char *szPath)
{
	// Look for file in current directory *first*
	if(GetFileAttributes(szFileName) != INVALID_FILE_ATTRIBUTES)
	{
		lstrcpy(szPath, szFileName);
	}
	else
	{
		// look for Debug\filename  or Release\filename
		char temp[MAX_PATH];
#ifdef _DEBUG
		wsprintf(temp, "Debug\\%s", szFileName);
#else
		wsprintf(temp, "Release\\%s", szFileName);
#endif

		if(GetFileAttributes(temp))
			lstrcpy(szPath, temp);
		else
			lstrcpy(szPath, "");
	}
}*/

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hwndList;
	char ach[80];

	switch(msg)
	{
	case WM_CREATE:
		hwndList = CreateWindowEx(WS_EX_CLIENTEDGE, "LISTBOX", 0, WS_CHILD|WS_VISIBLE
			|LBS_NOINTEGRALHEIGHT,
			0,0,0,0,hwnd,0,0,0);
		return 0;

	case WM_CLOSE:
		DestroyWindow(hwnd);
		return 0;

	case WM_SIZE:
		MoveWindow(hwndList, 0, 0, LOWORD(lParam), HIWORD(lParam), TRUE);
		return 0;

	case WM_DESTROY:
		PostQuitMessage(0);
		return 0;

	case WM_COMMAND:

		switch(LOWORD(wParam))
		{
		case 12345: case 54321:
			sprintf(ach, "Button %d pressed on window %x", LOWORD(wParam), lParam);
			SendMessage(hwndList, LB_ADDSTRING, 0, (LPARAM)ach);
			return 0;

		}

		return 0;

	}

	return DefWindowProc(hwnd, msg, wParam, lParam);
}

void gui()
{
	MSG msg;
	HWND hwnd;
	WNDCLASSEX wc = { sizeof(wc) };

	wc.hCursor = LoadCursor(0, MAKEINTRESOURCE(IDC_ARROW));
	wc.hInstance = GetModuleHandle(0);
	wc.lpfnWndProc = WndProc;
	wc.lpszClassName = "GlobarTest";

	RegisterClassEx(&wc);

	hwnd = CreateWindowEx(0, "GlobarTest", "GlobarTest", WS_OVERLAPPEDWINDOW|WS_VISIBLE, 
		CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT, 0, 0, GetModuleHandle(0), 0);
	
	// add the buttons! the Command-IDs will be sent to the window
	// we just created using a WM_COMMAND. The bitmaps displayed will use the bitmap
	// resources stored inside globar.dll, so we need to #include "resource.h" from
	// the DLL project.
	AddGlobalButton(hwnd, 12345, IDB_BITMAP1);
	AddGlobalButton(hwnd, 54321, IDB_BITMAP2);

	DisableByName("emule.exe");

	EnableGlobalButtons(TRUE);

	//pfnEnableButtons(hwnd);
	while(GetMessage(&msg, 0, 0, 0) > 0)
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	
	EnableGlobalButtons(FALSE);	
}

//
//	program entrypoint usage:
//
//	wincap <hwnd> [bitdepth]
//
int WINAPI WinMain(HINSTANCE hInst, HINSTANCE hPrev, LPSTR lpCmdLine, int nShowCmd)
{
	gui();
	return 0;
}
