#pragma once
#include <Windows.h>

void move_window_to_new_desktop(HWND hwnd);
void move_window_to_primary_desktop(HWND hwnd);
int GetCurrentDesktopGUIDIndexForWindow(HWND hwnd);
