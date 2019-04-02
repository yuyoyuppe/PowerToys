#ifndef CAPTIONBUTTON_INCLUDED
#define CAPTIONBUTTON_INCLUDED

#ifdef __cplusplus
extern "C" {
#endif

#define MAX_TITLE_BUTTONS	8


BOOL WINAPI Caption_InsertButton(HWND hwnd, UINT uCmd, int nBorder, HBITMAP hBmp);
BOOL WINAPI Caption_RemoveAllButtons(HWND hwnd);

#ifdef __cplusplus
}
#endif

#endif