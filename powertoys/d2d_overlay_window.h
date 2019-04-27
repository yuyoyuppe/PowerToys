#pragma once
#include "d2d_svg.h"
#include "d2d_window.h"
#include "monitors.h"

struct ScaleResult {
  double scale;
  RECT rect;
};

class D2DOverlaySVG : public D2DSVG {
public:
  D2DOverlaySVG& load(const std::wstring& filename, ID2D1DeviceContext5* d2d_dc);
  D2DOverlaySVG& resize(int x, int y, int width, int height, float fill, float max_scale = -1.0f);
  D2DOverlaySVG& find_thumbnail(const std::wstring& id);
  D2DOverlaySVG& find_window_group(const std::wstring& id);
  ScaleResult get_thumbnail_rect_and_scale(int x_offset, int y_offset, int window_cx, int window_cy, float fill);
  D2DOverlaySVG& toggle_window_group(bool active);
private:
  D2D1_POINT_2F thumbnail_top_left, thumbnail_bottom_right;
  RECT thumbnail_scaled_rect;
  winrt::com_ptr<ID2D1SvgElement> window_group;
};

class D2DOverlayWindow : public D2DWindow {
public:
  D2DOverlayWindow();
  void show(HWND active_window);
private:
  bool show_thumbnail(const ScaleResult& rect_and_scale);
  virtual void init() override;
  virtual void resize() override;
  virtual void render(ID2D1DeviceContext5* d2d_dc) override;
  virtual void on_show() override;
  virtual void on_hide() override;

  std::vector<MonitorInfo> monitors;
  MonitorInfo total_monitor;
  int monitor_dx = 0, monitor_dy = 0;
  WindowsColors colors;
  Animation animation;
  RECT window_rect = {};
  Tasklist tasklist;
  HTHUMBNAIL thumbnail;
  HWND active_window = nullptr;
  D2DOverlaySVG landscape, portrait;
  D2DOverlaySVG* use_overlay;
  D2DSVG no_active;
  std::vector<D2DSVG> arrows;
};
