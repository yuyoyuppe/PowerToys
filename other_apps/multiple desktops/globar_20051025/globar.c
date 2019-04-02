//
//	globar.c
//
//	GLObal titleBar library
//
//	www.catch22.net
//

#include <windows.h>
#include <psapi.h>
#include "resource.h"
#include "titlebar.h"

VOID InitThemeSupport();

//
// Export the API
//
#pragma comment(linker, "/opt:nowin98")
/*#pragma comment(linker, "/export:SetNotifyWindow=_SetNotifyWindow@4")
#pragma comment(linker, "/export:EnableGlobalButtons=_EnableGlobalButtons@4")
#pragma comment(linker, "/export:AddGlobalButton=_AddGlobalButton@12")
#pragma comment(linker, "/export:RemoveGlobalButton=_RemoveGlobalButton@4")
#pragma comment(linker, "/export:DisableByNameW=_DisableByNameW@4")
#pragma comment(linker, "/export:DisableByNameA=_DisableByNameA@4")

#pragma comment(linker, "/export:EnableButtons=_EnableButtons@4")*/

// needed for GetModuleBaseName
#pragma comment(lib, "psapi.lib")

HINSTANCE g_hInstance;
#define MAX_WINDOWS 256
#define MAX_BLACKLIST 100

//
//	Create a shared data section for the hook to live in
//
#pragma data_seg(".shared")
#pragma comment(linker, "/section:.shared,rws")

HHOOK   g_hHook			= 0;
HWND	g_hwndNotify	= 0;
BOOL	g_FirstInstance = FALSE;

UINT	g_nButtonCmdId[MAX_TITLE_BUTTONS] = { 0 };
UINT	g_nButtonResId[MAX_TITLE_BUTTONS] = { 0 };
int		g_nButtonCount  = 0;

WCHAR	g_szBlackList[MAX_BLACKLIST][64] = { 1 };
LONG	g_nBlackListLen = 0;

#pragma data_seg()

//
//	Use a *private* list to keep track of all windows that
//	the current process has subclassed - because we must remove
//	these subclasses when we unload
//
HWND			 g_hwndList[MAX_WINDOWS];
LONG			 g_nListLength = 0;
CRITICAL_SECTION g_csListLock;

BOOL WINAPI EnableButtons(HWND hwnd)
{
	LONG i;
	LONG style   = GetWindowLong(hwnd, GWL_STYLE);
	LONG desired = WS_OVERLAPPEDWINDOW;

	// check window is top-level
	//if(GetParent(hwnd) != 0 || 
	//	(GetWindowLong(hwnd, GWL_STYLE) & WS_CHILD) != 0)
	//	return FALSE;

	if((style & desired) != desired)
		return FALSE;

	EnterCriticalSection(&g_csListLock);

	if(g_nListLength < MAX_WINDOWS)
	{
		for(i = 0; i < g_nListLength; i++)
		{
			if(g_hwndList[i] == hwnd)
				break;
		}

		// if we reached the end of the list we didnt find the window,
		// so add those buttons!
		if(i == g_nListLength && g_nButtonCount > 0)
		{
			g_hwndList[g_nListLength++] = hwnd;

			for(i = 0; i < g_nButtonCount; i++)
			{
				HBITMAP hBitmap;
				hBitmap = LoadBitmap(g_hInstance, MAKEINTRESOURCE(g_nButtonResId[i]));
				
				Caption_InsertButton(hwnd, g_nButtonCmdId[i], 2, hBitmap);
			}
		}
	}

	LeaveCriticalSection(&g_csListLock);

	return TRUE;
}

BOOL WINAPI RemoveButtons(HWND hwnd)
{
	LONG i;

	EnterCriticalSection(&g_csListLock);

	for(i = 0; i < g_nListLength; i++)
	{
		if(g_hwndList[i] == hwnd)
		{
			g_hwndList[i] = g_hwndList[--g_nListLength];
			
			Caption_RemoveAllButtons(hwnd);

			break;
		}
	}

	LeaveCriticalSection(&g_csListLock);

	return TRUE;
}

BOOL RemoveAll()
{
	LONG i;

	EnterCriticalSection(&g_csListLock);

	for(i = 0; i < g_nListLength; i++)
	{
		Caption_RemoveAllButtons(g_hwndList[i]);
	}

	g_nListLength = 0;

	LeaveCriticalSection(&g_csListLock);
	return TRUE;
}


//
//	This is the system-wide hook
//
LRESULT CALLBACK CBTProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	if(nCode < 0)
		return CallNextHookEx(g_hHook, nCode, wParam, lParam);

	switch(nCode)
	{
	case HCBT_ACTIVATE:
		EnableButtons((HWND)wParam);
		break;

	// a window has been created. add the buttons!
	case HCBT_CREATEWND:
		//AddButtons((HWND)wParam);
		break;

	case HCBT_DESTROYWND:
		RemoveButtons((HWND)wParam);
		break;

	default:
		break;
	}

	return CallNextHookEx(g_hHook, nCode, wParam, lParam);
}

__declspec(dllexport)
BOOL WINAPI SetNotifyWindow(HWND hwnd)
{
	g_hwndNotify = hwnd;
	return TRUE;
}

__declspec(dllexport)
BOOL WINAPI AddGlobalButton(HWND hwndNotify, UINT nCmdId, int nBitmapId)
{
	if(g_FirstInstance != 1)
		return FALSE;

	if(g_nButtonCount < MAX_TITLE_BUTTONS)
	{
		g_hwndNotify   = hwndNotify;
		
		g_nButtonCmdId[g_nButtonCount] = nCmdId;
		g_nButtonResId[g_nButtonCount] = nBitmapId;

		g_nButtonCount++;

		return TRUE;
	}
	else
	{
		return FALSE;
	}
}

__declspec(dllexport)
BOOL WINAPI RemoveGlobalButton(UINT nCmdId)
{
	int i;

	if(g_FirstInstance != 1)
		return FALSE;

	for(i = 0; i < g_nButtonCount; i++)
	{
		if(g_nButtonCmdId[i] == nCmdId)
		{
			g_nButtonCmdId[i] = g_nButtonCmdId[g_nButtonCount-1];
			g_nButtonResId[i] = g_nButtonResId[g_nButtonCount-1];

			g_nButtonCount--;
			return TRUE;
		}
	}

	return FALSE;
}

//
//
//
__declspec(dllexport)
BOOL WINAPI EnableGlobalButtons(BOOL fEnable)
{
	if(g_FirstInstance != 1)
		return FALSE;

	if(g_hHook == 0 && fEnable == TRUE)
	{
		g_hHook = SetWindowsHookEx(WH_CBT, CBTProc, g_hInstance, 0);

		return g_hHook ? TRUE : FALSE;
	}
	else if(g_hHook != 0 && fEnable == FALSE)
	{
		return UnhookWindowsHookEx(g_hHook);
	}
	else
	{
		return FALSE;
	}
}

//
//	Disable buttons for specified exe
//
__declspec(dllexport)
BOOL WINAPI DisableByNameW(WCHAR *szName)
{
	if(g_nBlackListLen < MAX_BLACKLIST)
	{
		LONG i = g_nBlackListLen++;
	
		wcscpy(g_szBlackList[i], szName);
	}

	return FALSE;
}

//
//	Disable buttons for specified exe
//
__declspec(dllexport)
BOOL WINAPI DisableByNameA(char *szName)
{
	WCHAR wszName[MAX_PATH];

	MultiByteToWideChar(CP_ACP, 0, szName, -1, wszName, MAX_PATH);
	
	return DisableByNameW(wszName);
}

VOID ButtonPressed(UINT nCommandId, HWND hWnd)
{
	PostMessage(g_hwndNotify, WM_NULL, (WPARAM)nCommandId, (LPARAM)hWnd);
}

//
//	Inspect the application "blacklist" to decide if we want
//  to modify this app or not
//
BOOL CanCustomizeThisApp()
{
	WCHAR szBase[MAX_PATH];
	LONG i;
	
	GetModuleBaseNameW(GetCurrentProcess(), 0, szBase, MAX_PATH);
	
	for(i = 0; i < g_nBlackListLen; i++)
	{
		if(lstrcmpiW(szBase, g_szBlackList[i]) == 0)
			return FALSE;
	}
	
	return TRUE;
}

//
//	Main DLL entrypoint
//
BOOL CALLBACK DllMain(HINSTANCE hInst, DWORD dwReason, LPVOID lpReserved)
{
//	char ach[80];

	switch(dwReason)
	{
	case DLL_PROCESS_ATTACH:
		DisableThreadLibraryCalls(hInst);
		
		g_hInstance = hInst;
		g_FirstInstance++;

		if(!CanCustomizeThisApp())
			return FALSE;

		InitializeCriticalSection(&g_csListLock);

#ifdef _DEBUG
		wsprintf(ach, "Injected into %d\n", GetCurrentProcessId());
		OutputDebugString(ach);
#endif

		InitThemeSupport();
		break;

	case DLL_PROCESS_DETACH:

		RemoveAll();
		DeleteCriticalSection(&g_csListLock);

#ifdef _DEBUG		
		wsprintf(ach, "Unloading from %d\n", GetCurrentProcessId());
		OutputDebugString(ach);
#endif

		break;
	}

	return TRUE;
}
