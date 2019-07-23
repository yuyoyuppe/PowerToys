#pragma once
#include <functional>
#include <Windows.h>
/*
  In: HWND hwnd - of the window on which mouse is hovering
      RECT buttons - rect with the positions of the buttons
      RECT monitor - rect with the monitor to clip the popup window to
  Out: Rect of the created popup window, mouse out will be signalled when the
       mouse leaves it.
*/
typedef std::function<RECT(HWND hwnd, RECT buttons, POINT mouse_pos)> MouseInProc;
typedef std::function<void()> MouseOutProc;


void start_mouse_watcher(int ms_delay, int probe_ms_delay, MouseInProc on_mouse_in, MouseOutProc on_mouse_out, HWND popup);
void stop_mouse_watcher();
void update_mousein_wait(int ms_delay);
