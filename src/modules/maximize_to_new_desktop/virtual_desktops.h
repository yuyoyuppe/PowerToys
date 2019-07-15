#pragma once
#include <Windows.h>

void move_window_to_new_desktop(HWND hwnd);
void move_window_to_primary_desktop(HWND hwnd, bool close_desktop_if_last_window);
int get_desktop_index_for_window(HWND hwnd);
