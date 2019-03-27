#pragma once
#include <optional>
#include <Windows.h>

std::optional<RECT> get_maximize_button_pos(HWND hwnd);
std::optional<RECT> get_window_pos(HWND hwnd);
int run_message_loop();
