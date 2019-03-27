#pragma once
#include <optional>
#include <Windows.h>

std::optional<RECT> get_maximize_button(HWND hwnd);
int run_message_loop();
