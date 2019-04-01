#include "move_window.h"
#include <vector>
#include <algorithm>
#include <tuple>
#include <optional>

namespace {
  struct MonitorInfo {
    explicit MonitorInfo(RECT rect) : rect(rect) {}
    RECT rect;
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
  bool operator==(const MonitorInfo& lhs, const MonitorInfo& rhs) {
    auto lhs_tuple = std::make_tuple(lhs.rect.left, lhs.rect.right, lhs.rect.top, lhs.rect.bottom);
    auto rhs_tuple = std::make_tuple(rhs.rect.left, rhs.rect.right, rhs.rect.top, rhs.rect.bottom);
    return lhs_tuple == rhs_tuple;
  }
  BOOL CALLBACK monitor_enum_cb(HMONITOR monitor, HDC hdc, LPRECT rect, LPARAM data) {
    MONITORINFOEX monitor_info;
    monitor_info.cbSize = sizeof(MONITORINFOEX);
    GetMonitorInfo(monitor, &monitor_info);
    reinterpret_cast<std::vector<MonitorInfo>*>(data)->emplace_back(monitor_info.rcWork);
    return true;
  };
  // Returns monitor rects ordered from left to right
  std::vector<MonitorInfo> get_monitors() {
    std::vector<MonitorInfo> monitors;
    EnumDisplayMonitors(NULL, NULL, monitor_enum_cb, reinterpret_cast<LPARAM>(&monitors));
    std::sort(begin(monitors), end(monitors), [](const MonitorInfo& lhs, const MonitorInfo& rhs) {
      auto lhs_tuple = std::make_tuple(lhs.rect.left, lhs.rect.right, lhs.rect.top, lhs.rect.bottom);
      auto rhs_tuple = std::make_tuple(rhs.rect.left, rhs.rect.right, rhs.rect.top, rhs.rect.bottom);
      return lhs_tuple < rhs_tuple;
    });
    return monitors;
  }

  RECT get_window_monitor_rect(HWND hwnd) {
    auto monitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
    MONITORINFOEX monitor_info;
    monitor_info.cbSize = sizeof(MONITORINFOEX);
    GetMonitorInfo(monitor, &monitor_info);
    return monitor_info.rcWork;
  }

  int get_monitor_index(const std::vector<MonitorInfo>& monitors, RECT monitor_rect) {
    auto iter = std::find(begin(monitors), end(monitors), MonitorInfo(monitor_rect));
    if (iter != end(monitors))
      return iter - begin(monitors);
    // If not found return the middle monitor
    return monitors.size() / 2;
  }

  RECT translate_monitors(RECT window, RECT src_monitor, RECT dest_monitor) {
    auto src_width = src_monitor.right - src_monitor.left;
    auto src_heigh = src_monitor.bottom - src_monitor.top;
    double left = double(window.left - src_monitor.left) / src_width;
    double right = double(window.right - src_monitor.left) / src_width;
    double top = double(window.top - src_monitor.top) / src_heigh;
    double bottom = double(window.bottom - src_monitor.top) / src_heigh;
    auto dest_width = dest_monitor.right - dest_monitor.left;
    auto dest_height = dest_monitor.bottom - dest_monitor.top;
    return {
      (int)(left * dest_width) + dest_monitor.left,
      (int)(top * dest_height) + dest_monitor.top,
      (int)(right * dest_width) + dest_monitor.left,
      (int)(bottom * dest_height) + dest_monitor.top
    };
  }
}

namespace move_window {
  void maximize(HWND hwnd) {}
  void restore(HWND hwnd) {}
  void minimize(HWND hwnd) {}

  void snap_top_left(HWND hwnd) {}
  void snap_top(HWND hwnd) {}
  void snap_top_right(HWND hwnd) {}
  void snap_left(HWND hwnd) {}
  void snap_right(HWND hwnd) {}
  void snap_bottom_left(HWND hwnd) {}
  void snap_bottom(HWND hwnd) {}
  void snap_bottom_right(HWND hwnd) {}

  void to_screen_left(HWND hwnd) {
    auto monitors = get_monitors();
    auto current_monitor_rect = get_window_monitor_rect(hwnd);
    auto current = get_monitor_index(monitors, current_monitor_rect);
    auto left = current - 1;
    if (left == -1)
      left = monitors.size() - 1;
    RECT window_pos;
    if (GetWindowRect(hwnd, &window_pos) == 0) {
      window_pos = current_monitor_rect;
    }
    auto new_pos = translate_monitors(window_pos, current_monitor_rect, monitors[left].rect);
    //auto new_pos = window_pos;
    SetWindowPos(hwnd, HWND_TOP, new_pos.left, new_pos.top, new_pos.right - new_pos.left, new_pos.bottom - new_pos.top, SWP_NOREDRAW);
  }
  void to_screen_right(HWND hwnd) {
    auto monitors = get_monitors();
    auto current_monitor_rect = get_window_monitor_rect(hwnd);
    auto current = get_monitor_index(monitors, current_monitor_rect);
    auto right = current + 1;
    if (right == monitors.size())
      right = 0;
    RECT window_pos;
    if (GetWindowRect(hwnd, &window_pos) == 0) {
      window_pos = current_monitor_rect;
    }
    auto new_pos = translate_monitors(window_pos, current_monitor_rect, monitors[right].rect);
    SetWindowPos(hwnd, HWND_TOP, new_pos.left, new_pos.top, new_pos.right - new_pos.left, new_pos.bottom - new_pos.top, SWP_NOREDRAW);
  }

  void to_next_desktop(HWND hwnd) {}
  void to_prev_desktop(HWND hwnd) {}
}
