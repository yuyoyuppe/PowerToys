#pragma once
#include <Windows.h>
#include <vector>

struct MonitorInfo {
  explicit MonitorInfo(RECT rect) : rect(rect) {}
  RECT rect;
  int left() const { return rect.left; }
  int right() const { return rect.right; }
  int top() const { return rect.top; }
  int bottom() const { return rect.bottom; }
  int height() const { return rect.bottom - rect.top; };
  int width() const  { return rect.right - rect.left; };
  POINT top_left() const     { return { rect.left,               rect.top }; };
  POINT top_middle() const   { return { rect.left + width() / 2, rect.top }; };
  POINT top_right() const    { return { rect.right,              rect.top }; };
  POINT middle_left() const  { return { rect.left,               rect.top + height() / 2 }; };
  POINT middle() const       { return { rect.left + width() / 2, rect.top + height() / 2 }; };
  POINT middle_right() const { return { rect.right,              rect.top + height() / 2 }; };
  POINT bottom_left() const  { return { rect.left,               rect.bottom }; };
  POINT bottm_midle() const  { return { rect.left + width() / 2, rect.bottom }; };
  POINT bottom_right() const { return { rect.right,              rect.bottom }; };
};

bool operator==(const MonitorInfo& lhs, const MonitorInfo& rhs);


// Returns monitor rects ordered from left to right
std::vector<MonitorInfo> get_monitors(bool include_toolbar);
// Return primary display
MonitorInfo get_primary_monitor();
// Return monitor on which hwnd window is displayed
MonitorInfo get_window_monitor(HWND hwnd);
// Return monitor nearest to a point
MonitorInfo get_point_monitor(POINT p);
// Return monitor info given a HMONITOR
MonitorInfo get_monitor_info(HMONITOR monitor);
// Gets numeric index of the given monitor, for easy access to the screens on the left and right
int get_monitor_index(const std::vector<MonitorInfo>& monitors, const MonitorInfo& monitor);
// Scales window from src_monitor coordinates to des_monitor coordinates. Returned RECT will show the window
// in the same proportion on the dest monitor as it was on the src one.
RECT translate_monitors(const RECT& window, const MonitorInfo& src_monitor, const MonitorInfo& dest_monitor);
