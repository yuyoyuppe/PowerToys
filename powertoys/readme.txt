main.cpp:

Calling SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE) before any other
WINAPI call disables Windows scaling, leaving everything to us.

popup_window:

Class for creating and displaying windows. Used params:
  WS_EX_TOOLWINDOW - window wont appear in Alt-Tab and on the taskbar
  WS_EX_TOPMOST - window should be on top
  WS_EX_LAYERED - can be made transparent by a call to SetLayeredWindowAttributes
  WS_EX_TRANSPARENT - window will be "clickthrough", all clicks will go to the
                      underlying window
  WS_POPUP - minimal window

We keep hash map of WM_PAINT handlers for each window (paint_procedures). This
way we can easily substitute paint handler to customize window content. We can
even do that when resizing the window.

keyboard_watcher:

Uses SetWindowsHookEx to install system-wide hook that intercepts keyboard
events. After WinKey is pressed, conditional variable is signaled and
held_delay_thread_proc thread waits for specified amount of time. If the key
is still pressed on_held callback is called.

winkey_handler:

GetForegroundWindow returns current active window. Right now no filtering is done,
so for example HWND of the start menu itself will be returned if it is visible.
GetDpiForWindow gets the window DPI which we will use for scaling. After
displaying the window we call SetForegroundWindow, to return focus to the window
that was originally on top.

