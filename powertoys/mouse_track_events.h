#pragma once
class MouseTrackEvents
{
public:
  MouseTrackEvents();
  void OnMouseMove(HWND hwnd, DWORD dwFlags);
  void Reset(HWND hwnd);
private:
  bool m_bMouseTracking;
};
