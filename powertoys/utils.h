#pragma once
#include <optional>
#include <Windows.h>

std::optional<RECT> get_button_pos(HWND hwnd);
std::optional<RECT> get_window_pos(HWND hwnd);
std::optional<POINT> get_mouse_pos();

int run_message_loop();
