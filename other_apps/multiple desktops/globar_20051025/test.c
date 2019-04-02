//
//	test.c
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
//	BOOL WINAPI EnableGlobalButtons(BOOL fEnable)
//
//
#include <windows.h>
#include <stdio.h>
#include "resource.h"

#pragma comment(linker, "/OPT:NOWIN98")

#define ENABLEBUTTONAPI "EnableGlobalButtons"
#define ADDBUTTONAPI	"AddGlobalButton"
#define GLOBARDLL		"globar.dll"

typedef BOOL (WINAPI * ENABLEPROC)(BOOL);
typedef BOOL (WINAPI * ADDPROC)(HWND, UINT, UINT);

__declspec(dllimport) HWND WINAPI GetConsoleWindow();

//
//	Silly function to find the capture.dll file - in visual studio,
//	it is always in .\debug or .\release and not in the current directory
//	so this is just a bodge to make building+testing easier
//
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
}

//
//	program entrypoint usage:
//
//	wincap <hwnd> [bitdepth]
//
int main(int argc, char *argv[])
{
	char    szModule[MAX_PATH] = GLOBARDLL;
	char	szFullPath[MAX_PATH];
	char   *fp;

	HMODULE				hModule;
	ENABLEPROC			pfnEnableGlobalButtons;
	ADDPROC				pfnAddGlobalButton;

	//
	// get path to the capture.dll module
	//
	FindModuleFileName(GLOBARDLL, szModule);
	
	if(!GetFullPathName(szModule, MAX_PATH, szFullPath, &fp))
	{
		printf("Failed to load %s\n", GLOBARDLL);
		return 1;
	}

	//
	//	Load GLOBAR.DLL
	//
	if((hModule = LoadLibrary(szFullPath)) == 0)
	{
		printf("Failed to load %s\n", szFullPath);
		return 1;
	}

	//
	//	Get the CAPTURE.DLL exported functions
	//
	pfnEnableGlobalButtons  = (ENABLEPROC)GetProcAddress(hModule, ENABLEBUTTONAPI);
	pfnAddGlobalButton      = (ADDPROC)GetProcAddress(hModule, ADDBUTTONAPI);
	
	if(pfnEnableGlobalButtons != 0 && pfnAddGlobalButton != 0)
	{
		MSG msg;
		HWND hwnd = (HWND)0x00020666;	//GetConsoleWindow()

		pfnAddGlobalButton(hwnd, 12345, IDB_BITMAP1);
		pfnAddGlobalButton(hwnd, 54321, IDB_BITMAP2);
		pfnEnableGlobalButtons(TRUE);

		while(GetMessage(&msg, 0, 0, 0) > 0)
		{
			printf("smeg\n");
		}
		
		getch();

		pfnEnableGlobalButtons(FALSE);
	}
	else
	{
		printf("Failed to load %s [%x]\n", szFullPath, GetLastError());
	}

	FreeLibrary(hModule);

	return 0;
}
