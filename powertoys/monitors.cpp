#include "monitors.h"
#include <tuple>
#include <algorithm>

bool operator==(const MonitorInfo& lhs, const MonitorInfo& rhs) {
  auto lhs_tuple = std::make_tuple(lhs.rect.left, lhs.rect.right, lhs.rect.top, lhs.rect.bottom);
  auto rhs_tuple = std::make_tuple(rhs.rect.left, rhs.rect.right, rhs.rect.top, rhs.rect.bottom);
  return lhs_tuple == rhs_tuple;
}

static BOOL CALLBACK get_displays_enum_cb(HMONITOR monitor, HDC hdc, LPRECT rect, LPARAM data) {
  MONITORINFOEX monitor_info;
  monitor_info.cbSize = sizeof(MONITORINFOEX);
  GetMonitorInfo(monitor, &monitor_info);
  reinterpret_cast<std::vector<MonitorInfo>*>(data)->emplace_back(monitor_info.rcWork);
  return true;
};

std::vector<MonitorInfo> get_monitors() {
  std::vector<MonitorInfo> monitors;
  EnumDisplayMonitors(NULL, NULL, get_displays_enum_cb, reinterpret_cast<LPARAM>(&monitors));
  std::sort(begin(monitors), end(monitors), [](const MonitorInfo& lhs, const MonitorInfo& rhs) {
    auto lhs_tuple = std::make_tuple(lhs.rect.left, lhs.rect.right, lhs.rect.top, lhs.rect.bottom);
    auto rhs_tuple = std::make_tuple(rhs.rect.left, rhs.rect.right, rhs.rect.top, rhs.rect.bottom);
    return lhs_tuple < rhs_tuple;
  });
  return monitors;
}

static BOOL CALLBACK get_primary_display_enum_cb(HMONITOR monitor, HDC hdc, LPRECT rect, LPARAM data) {
  MONITORINFOEX monitor_info;
  monitor_info.cbSize = sizeof(MONITORINFOEX);
  GetMonitorInfo(monitor, &monitor_info);
  if (monitor_info.dwFlags & MONITORINFOF_PRIMARY)
    reinterpret_cast<MonitorInfo*>(data)->rect = monitor_info.rcWork;
  return true;
};

MonitorInfo get_primary_monitor() {
  MonitorInfo primary({});
  EnumDisplayMonitors(NULL, NULL, get_primary_display_enum_cb, reinterpret_cast<LPARAM>(&primary));
  return primary;
}

MonitorInfo get_window_monitor(HWND hwnd) {
  auto monitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
  MONITORINFOEX monitor_info;
  monitor_info.cbSize = sizeof(MONITORINFOEX);
  GetMonitorInfo(monitor, &monitor_info);
  return MonitorInfo(monitor_info.rcWork);
}

int get_monitor_index(const std::vector<MonitorInfo>& monitors, const MonitorInfo& monitor) {
  auto iter = std::find(begin(monitors), end(monitors), monitor);
  if (iter != end(monitors))
    return iter - begin(monitors);
  // If not found return the middle monitor
  return monitors.size() / 2;
}

RECT translate_monitors(const RECT& window, const MonitorInfo& src_monitor, const MonitorInfo& dest_monitor) {
  auto src_width = src_monitor.width();
  auto src_heigh = src_monitor.height();
  double left = double(window.left - src_monitor.left()) / src_width;
  double right = double(window.right - src_monitor.left()) / src_width;
  double top = double(window.top - src_monitor.top()) / src_heigh;
  double bottom = double(window.bottom - src_monitor.top()) / src_heigh;
  auto dest_width = dest_monitor.width();
  auto dest_height = dest_monitor.height();
  return {
    (int)(left * dest_width) + dest_monitor.left(),
    (int)(top * dest_height) + dest_monitor.top(),
    (int)(right * dest_width) + dest_monitor.left(),
    (int)(bottom * dest_height) + dest_monitor.top()
  };
}
