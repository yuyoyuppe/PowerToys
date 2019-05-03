#include "pch.h"
#include "mouse_track_events.h"

MouseTrackEvents::MouseTrackEvents() : m_bMouseTracking(false) {
}

void MouseTrackEvents::OnMouseMove(HWND hwnd, DWORD dwFlags) {
  if (!m_bMouseTracking)
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
    m_bMouseTracking = true;
  }
}

void MouseTrackEvents::Reset(HWND hwnd)
{
  m_bMouseTracking = false;
}
