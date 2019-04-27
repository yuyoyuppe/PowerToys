#pragma once
#include <optional>
#include <Windows.h>

// Returns RECT with positions of the minmize/maximize buttons of the given window.
// Does not always work, since some apps draw custom toolbars.
std::optional<RECT> get_button_pos(HWND hwnd);
// Gets position of given window.
std::optional<RECT> get_window_pos(HWND hwnd);
// Gets mouse postion.
std::optional<POINT> get_mouse_pos();
// Compare rects
bool operator==(const RECT& lhs, const RECT& rhs);
bool operator!=(const RECT& lhs, const RECT& rhs);
bool operator<(const RECT& lhs, const RECT& rhs);
// Moves and/or resizes small_rect to fit inside big_rect.
RECT keep_rect_inside_rect(const RECT& small_rect, const RECT& big_rect);
// Initializes and runs windows message loop
int run_message_loop();
void ShowLastErrorMessage(LPTSTR lpszFunction, DWORD dw);
