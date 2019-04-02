#ifndef GLOBAR_INCLUDED
#define GLOBAR_INCLUDED

__declspec(dllimport)
BOOL WINAPI AddGlobalButton(HWND hwndNotify, UINT nCmdId, int nBitmapId);

__declspec(dllimport)
BOOL WINAPI RemoveGlobalButton(UINT nCmdId);

__declspec(dllimport)
BOOL WINAPI EnableGlobalButtons(BOOL fEnable);

__declspec(dllimport)
BOOL WINAPI DisableByNameA(char *szName);

__declspec(dllimport)
BOOL WINAPI DisableByNameW(WCHAR *szName);

#ifdef UNICODE
#define DisableByName DisableByNameW
#else
#define DisableByName DisableByNameA
#endif

#endif