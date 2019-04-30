#pragma once
#include "d2d_svg.h"
#include "d2d_window.h"
#include "d2d_text.h"
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
  winrt::com_ptr<ID2D1SvgElement> find_element(const std::wstring& id);
  const float xoff = 0.009;
  D2D1_RECT_F get_maximize_label() {
    D2D1_RECT_F result;
    float height = thumbnail_scaled_rect.bottom - thumbnail_scaled_rect.top;
    float width = thumbnail_scaled_rect.right - thumbnail_scaled_rect.left;
    result.top = thumbnail_scaled_rect.bottom + height * 0.21;
    result.bottom = thumbnail_scaled_rect.bottom + height * 0.31;
    result.left = thumbnail_scaled_rect.left + width * xoff;
    result.right = thumbnail_scaled_rect.right + width * xoff;
    return result;
  }
  D2D1_RECT_F get_minimize_label() {
    D2D1_RECT_F result;
    float height = thumbnail_scaled_rect.bottom - thumbnail_scaled_rect.top;
    float width = thumbnail_scaled_rect.right - thumbnail_scaled_rect.left;
    result.top = thumbnail_scaled_rect.bottom + height * 0.8;
    result.bottom = thumbnail_scaled_rect.bottom + height * 0.9;
    result.left = thumbnail_scaled_rect.left + width * xoff;
    result.right = thumbnail_scaled_rect.right + width * xoff;
    return result;
  }
  D2D1_RECT_F get_snap_left() {
    D2D1_RECT_F result;
    float height = thumbnail_scaled_rect.bottom - thumbnail_scaled_rect.top;
    float width = thumbnail_scaled_rect.right - thumbnail_scaled_rect.left;
    result.top = thumbnail_scaled_rect.bottom + height * 0.5;
    result.bottom = thumbnail_scaled_rect.bottom + height * 0.6;
    result.left = thumbnail_scaled_rect.left + width * xoff;
    result.right = thumbnail_scaled_rect.left + width * 0.33 + width * xoff;
    return result;
  }
  D2D1_RECT_F get_snap_right() {
    D2D1_RECT_F result;
    float height = thumbnail_scaled_rect.bottom - thumbnail_scaled_rect.top;
    float width = thumbnail_scaled_rect.right - thumbnail_scaled_rect.left;
    result.top = thumbnail_scaled_rect.bottom + height * 0.5;
    result.bottom = thumbnail_scaled_rect.bottom + height * 0.6;
    result.left = thumbnail_scaled_rect.left + width * 0.67 + width * xoff;
    result.right = thumbnail_scaled_rect.right + width + width * xoff;
    return result;
  }
private:
  D2D1_POINT_2F thumbnail_top_left, thumbnail_bottom_right;
  RECT thumbnail_scaled_rect;
  winrt::com_ptr<ID2D1SvgElement> window_group;
};

struct AnimateKeys {
  Animation animation;
  D2D1_COLOR_F original;
  winrt::com_ptr<ID2D1SvgElement> button;
  int vk_code;
};

class D2DOverlayWindow : public D2DWindow {
public:
  D2DOverlayWindow();
  void show(HWND active_window);
  void animate(int vk_code);
private:
  void animate(int vk_code, int offset);
  bool show_thumbnail(const RECT& rect_and_scale);
  void hide_thumbnail();
  virtual void init() override;
  virtual void resize() override;
  virtual void render(ID2D1DeviceContext5* d2d_dc) override;
  virtual void on_show() override;
  virtual void on_hide() override;

  bool visible = false;
  std::vector<AnimateKeys> key_animations;
  std::vector<MonitorInfo> monitors;
  MonitorInfo total_monitor;
  int monitor_dx = 0, monitor_dy = 0;
  D2DText text;
  WindowsColors colors;
  Animation animation;
  RECT window_rect = {};
  Tasklist tasklist;
  std::vector<TasklistButton> tasklist_buttons;
  std::chrono::system_clock::time_point update_timestamp;
  HTHUMBNAIL thumbnail;
  HWND active_window = nullptr;
  D2DOverlaySVG landscape, portrait;
  D2DOverlaySVG* use_overlay;
  D2DSVG no_active;
  std::vector<D2DSVG> arrows;
};
