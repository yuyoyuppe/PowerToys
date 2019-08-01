#pragma once
#include <Windows.h>

void move_window_to_new_desktop(HWND hwnd, HWND popupWindow);
void move_window_to_primary_desktop(HWND hwnd, bool close_desktop_if_last_window, HWND popupWindow);
int get_desktop_index_for_window(HWND hwnd) noexcept;
