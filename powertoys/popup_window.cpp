#include "popup_window.h"

std::mutex PopupWindow::static_mutex;
bool PopupWindow::window_class_initialized;
std::unordered_map<HWND, PaintProc> PopupWindow::paint_proc;