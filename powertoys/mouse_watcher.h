#pragma once
#include <functional>
#include <Windows.h>
/*
  In: HWND - of the window on which mouse is hovering
      RECT - rect with the positions of the buttons
  Out: Rect of the created popup window, mouse out will be signalled when the
       mouse leaves it.
*/
typedef std::function<RECT(HWND, RECT)> MouseInProc;
typedef std::function<void()> MouseOutProc;


void start_mouse_watcher(int ms_delay, int probe_ms_delay, MouseInProc on_mouse_in, MouseOutProc on_mouse_out);
