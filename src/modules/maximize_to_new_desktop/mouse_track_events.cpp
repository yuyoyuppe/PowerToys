#include "pch.h"
#include "mouse_track_events.h"

MouseTrackEvents::MouseTrackEvents() : mouse_tracking(false) {
}

void MouseTrackEvents::on_moude_move(HWND hwnd, DWORD dwFlags) {
  if (!mouse_tracking)
  {
    // Enable mouse tracking.
    TRACKMOUSEEVENT tme;
    tme.cbSize = sizeof(tme);
    tme.hwndTrack = hwnd;
    tme.dwFlags = dwFlags;
    tme.dwHoverTime = HOVER_DEFAULT;
    if(!TrackMouseEvent(&tme)) {
      int a = 2;
    }
    mouse_tracking = true;
  }
}

void MouseTrackEvents::reset(HWND hwnd)
{
  mouse_tracking = false;
}
