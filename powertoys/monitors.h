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
std::vector<MonitorInfo> get_monitors();
MonitorInfo get_primary_monitor();
MonitorInfo get_window_monitor(HWND hwnd);
int get_monitor_index(const std::vector<MonitorInfo>& monitors, const MonitorInfo& monitor);
RECT translate_monitors(const RECT& window, const MonitorInfo& src_monitor, const MonitorInfo& dest_monitor);
