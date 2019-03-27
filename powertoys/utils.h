#pragma once
#include <optional>
#include <Windows.h>

std::optional<RECT> getWindowMaximizeButton(HWND hwnd);
