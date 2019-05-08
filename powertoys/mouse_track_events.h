#pragma once
class MouseTrackEvents
{
public:
  MouseTrackEvents();
  void on_moude_move(HWND hwnd, DWORD dwFlags);
  void reset(HWND hwnd);
private:
  bool mouse_tracking;
};
