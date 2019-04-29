#include "pch.h"
#include "d2d_overlay_window.h"
#include "monitors.h"
#include "tasklist_positions.h"
#include "keyboard_watcher.h"


D2DOverlaySVG& D2DOverlaySVG::load(const std::wstring& filename, ID2D1DeviceContext5* d2d_dc) {
  D2DSVG::load(filename, d2d_dc);
  window_group = nullptr;
  thumbnail_top_left = {};
  thumbnail_bottom_right = {};
  thumbnail_scaled_rect = {};
  return *this;
}

D2DOverlaySVG& D2DOverlaySVG::resize(int x, int y, int width, int height, float fill, float max_scale) {
  D2DSVG::resize(x, y, width, height, fill, max_scale);
  if (thumbnail_bottom_right.x != 0 && thumbnail_bottom_right.y != 0) {
    auto scaled_top_left = transform.TransformPoint(thumbnail_top_left);
    auto scanled_bottom_right = transform.TransformPoint(thumbnail_bottom_right);
    thumbnail_scaled_rect.left = (int)scaled_top_left.x;
    thumbnail_scaled_rect.top = (int)scaled_top_left.y;
    thumbnail_scaled_rect.right = (int)scanled_bottom_right.x;
    thumbnail_scaled_rect.bottom = (int)scanled_bottom_right.y;
  }
  return *this;
}

D2DOverlaySVG& D2DOverlaySVG::find_thumbnail(const std::wstring& id) {
  winrt::com_ptr<ID2D1SvgElement> thumbnail_box;
  winrt::check_hresult(svg->FindElementById(id.c_str(), thumbnail_box.put()));
  winrt::check_hresult(thumbnail_box->GetAttributeValue(L"x", &thumbnail_top_left.x));
  winrt::check_hresult(thumbnail_box->GetAttributeValue(L"y", &thumbnail_top_left.y));
  winrt::check_hresult(thumbnail_box->GetAttributeValue(L"width", &thumbnail_bottom_right.x));
  thumbnail_bottom_right.x += thumbnail_top_left.x;
  winrt::check_hresult(thumbnail_box->GetAttributeValue(L"height", &thumbnail_bottom_right.y));
  thumbnail_bottom_right.y += thumbnail_top_left.y;
  return *this;
}

D2DOverlaySVG& D2DOverlaySVG::find_window_group(const std::wstring& id) {
  window_group = nullptr;
  winrt::check_hresult(svg->FindElementById(id.c_str(), window_group.put()));
  return *this;
}

ScaleResult D2DOverlaySVG::get_thumbnail_rect_and_scale(int x_offset, int y_offset, int window_cx, int window_cy, float fill) {
  if (thumbnail_bottom_right.x == 0 && thumbnail_bottom_right.y == 0)
    return {};
  int thumbnail_scaled_rect_width = thumbnail_scaled_rect.right - thumbnail_scaled_rect.left;
  int thumbnail_scaled_rect_heigh = thumbnail_scaled_rect.bottom - thumbnail_scaled_rect.top;
  if (thumbnail_scaled_rect_heigh == 0 || thumbnail_scaled_rect_width == 0 ||
    window_cx == 0 || window_cy == 0) {
    return {};
  }
  float scale_h = fill * thumbnail_scaled_rect_width / window_cx;
  float scale_v = fill * thumbnail_scaled_rect_heigh / window_cy;
  float use_scale = min(scale_h, scale_v);
  RECT thumb_rect;
  thumb_rect.left = thumbnail_scaled_rect.left + (int)(thumbnail_scaled_rect_width - use_scale * window_cx) / 2 + x_offset;
  thumb_rect.right = thumbnail_scaled_rect.right - (int)(thumbnail_scaled_rect_width - use_scale * window_cx) / 2 + x_offset;
  thumb_rect.top = thumbnail_scaled_rect.top + (int)(thumbnail_scaled_rect_heigh - use_scale * window_cy) / 2 + y_offset;
  thumb_rect.bottom = thumbnail_scaled_rect.bottom - (int)(thumbnail_scaled_rect_heigh - use_scale * window_cy) / 2 + y_offset;
  ScaleResult result;
  result.scale = use_scale;
  result.rect = thumb_rect;
  return result;
}

D2DOverlaySVG& D2DOverlaySVG::toggle_window_group(bool active) {
  if (window_group)
    window_group->SetAttributeValue(L"fill-opacity", active ? 1.0f : 0.3f);
  return *this;
}

D2DOverlayWindow::D2DOverlayWindow() : animation(0.2), total_monitor({})
{ }

void D2DOverlayWindow::show(HWND active_window) {
  if (visible) {
    return;
  }
  visible = true;
  this->active_window = active_window;
  auto old_bck = colors.start_color_menu;
  if (initialized && colors.update()) {
    // update background colors
    landscape.recolor(old_bck, colors.start_color_menu);
    portrait.recolor(old_bck, colors.start_color_menu);
    for (auto& arrow : arrows) {
      arrow.recolor(old_bck, colors.start_color_menu);
    }
    if (colors.light_mode) {
      landscape.recolor(0xDDDDDD, 0x222222);
      portrait.recolor(0xDDDDDD, 0x222222);
      for (auto& arrow : arrows) {
        arrow.recolor(0xDDDDDD, 0x222222);
      }
    } else {
      landscape.recolor(0x222222, 0xDDDDDD);
      portrait.recolor(0x222222, 0xDDDDDD);
      for (auto& arrow : arrows) {
        arrow.recolor(0x222222, 0xDDDDDD);
      }
    }
  }
  monitors = get_monitors(true);
  // calculate the rect covering all the screens
  total_monitor = monitors[0];
  for (auto& monitor : monitors) {
    total_monitor.rect.left = min(total_monitor.rect.left, monitor.rect.left);
    total_monitor.rect.top = min(total_monitor.rect.top, monitor.rect.top);
    total_monitor.rect.right = max(total_monitor.rect.right, monitor.rect.right);
    total_monitor.rect.bottom = max(total_monitor.rect.bottom, monitor.rect.bottom);
  }
  // make sure top-right corner of all the monitor rects is (0,0)
  monitor_dx = -total_monitor.left();
  monitor_dy = -total_monitor.top();
  total_monitor.rect.left += monitor_dx;
  total_monitor.rect.right += monitor_dx;
  total_monitor.rect.top += monitor_dy;
  total_monitor.rect.bottom += monitor_dy;
  tasklist.update();
  if (active_window) {
    // Ignore errors, if this fails we will just not show the thumbnail
    DwmRegisterThumbnail(hwnd, active_window, &thumbnail);
  }
  animation.reset();
  auto primary_screen = get_primary_monitor();
  D2DWindow::show(primary_screen.left(), primary_screen.top(), primary_screen.width(), primary_screen.height());
}

void D2DOverlayWindow::on_show() { 
  // show override does everything
}

void D2DOverlayWindow::on_hide() {
  visible = false;
  if (thumbnail) {
    DwmUnregisterThumbnail(thumbnail);
  }
}

void D2DOverlayWindow::init() {
  colors.update();
  landscape.load(L"svgs\\overlay.svg", d2d_dc.get())
           .find_thumbnail(L"path-1")
           .find_window_group(L"Group-1")
           //.toggle_element(L"windows_bg", false)
           .recolor(0x000000, colors.start_color_menu);
  portrait.load(L"svgs\\overlay_portrait.svg", d2d_dc.get())
          .find_thumbnail(L"path-1")
          .find_window_group(L"Group-1")
          //.toggle_element(L"windows_bg", false)
          .recolor(0x000000, colors.start_color_menu);
  no_active.load(L"svgs\\no_active_window.svg", d2d_dc.get());
  arrows.resize(10);
  for (unsigned i = 0; i < arrows.size(); ++i) {
    arrows[i].load(L"svgs\\" + std::to_wstring((i + 1) % 10) + L".svg", d2d_dc.get())
             .recolor(0x000000, colors.start_color_menu);
  }
  if (!colors.light_mode) {
    landscape.recolor(0x222222, 0xDDDDDD);
    portrait.recolor(0x222222, 0xDDDDDD);
    for (auto& arrow : arrows) {
      arrow.recolor(0x222222, 0xDDDDDD);
    }
  }
}

void D2DOverlayWindow::resize() {
  window_rect = *get_window_pos(hwnd);
  float no_active_scale;
  if (window_width > window_height) {
    use_overlay = &landscape;
    no_active_scale = 0.3f;
  } else {
    use_overlay = &portrait;
    no_active_scale = 0.5f;
  }
  use_overlay->resize(0, 0, window_width, window_height, 0.8f);
  auto thumb_no_active_rect = use_overlay->get_thumbnail_rect_and_scale(0, 0, no_active.width(), no_active.height(), no_active_scale).rect;
  no_active.resize(thumb_no_active_rect.left,
                   thumb_no_active_rect.top,
                   thumb_no_active_rect.right - thumb_no_active_rect.left,
                   thumb_no_active_rect.bottom - thumb_no_active_rect.top,
                   1.0f);
}

void render_arrow(D2DSVG& arrow, TasklistButton& button, RECT window, float max_scale, ID2D1DeviceContext5* d2d_dc) {
  int dx = 0, dy = 0;
  // Calculate taskbar orientation
  arrow.toggle_element(L"left", false);
  arrow.toggle_element(L"right", false);
  arrow.toggle_element(L"top", false);
  arrow.toggle_element(L"bottom", false);
  if (button.x <= window.left) { // taskbar on left
    dx = 1;
    arrow.toggle_element(L"left", true);
  }
  if (button.x >= window.right) { // taskbar on right
    dx = -1;
    arrow.toggle_element(L"right", true);
  }
  if (button.y <= window.top) { // taskbar on top
    dy = 1;
    arrow.toggle_element(L"top", true);
  }
  if (button.y >= window.bottom) { // taskbar on bottom
    dy = -1;
    arrow.toggle_element(L"bottom", true);
  }
  double arrow_ratio = (double)arrow.height() / arrow.width();
  if (dy != 0) {
    // assume button is 25% wider than taller, +10% to make room for each of the arrows that are hidden
    auto render_arrow_width = (int)(button.height * 1.25f * 1.2f);
    auto render_arrow_height = (int)(render_arrow_width * arrow_ratio);
    auto y_edge = dy == -1 ? button.y : button.y + button.height;
    arrow.resize(button.x + (button.width - render_arrow_width) / 2,
                 dy == -1 ? button.y - render_arrow_height : 0,
                 render_arrow_width, render_arrow_height, 0.95f, max_scale)
         .render(d2d_dc);
  }
  else {
    // same as above - make room for the hidden arrow
    auto render_arrow_height = (int)(button.height * 1.2f);
    auto render_arrow_width = (int)(render_arrow_height / arrow_ratio);
    arrow.resize(dx == -1 ? button.x - render_arrow_width : 0,
                 button.y + (button.height - render_arrow_height) / 2,
                 render_arrow_width, render_arrow_height, 0.95f, max_scale)
         .render(d2d_dc);
  }
}

bool D2DOverlayWindow::show_thumbnail(const RECT& rect) {
  if (!thumbnail)
    return false;
  SIZE thumb_size;
  DWM_THUMBNAIL_PROPERTIES thumb_properties;
  thumb_properties.dwFlags = DWM_TNP_SOURCECLIENTAREAONLY | DWM_TNP_VISIBLE | DWM_TNP_RECTDESTINATION;
  thumb_properties.fSourceClientAreaOnly = FALSE;
  thumb_properties.fVisible = TRUE;
  thumb_properties.rcDestination = rect;
  if (DwmUpdateThumbnailProperties(thumbnail, &thumb_properties) != S_OK)
    return false;
  return true;
}

void D2DOverlayWindow::render(ID2D1DeviceContext5* d2d_dc) {
  if (!visible || !winkey_held()) {
    hide();
    return;
  }
  d2d_dc->Clear();
  auto tasklist_buttons = tasklist.get_buttons();
  int x_offset = 0, y_offset = 0;
  double anim_value = 1 - animation.value();
  if (!tasklist_buttons.empty()) {
    if (tasklist_buttons[0].x <= window_rect.left) { // taskbar on left
      x_offset = (int)(-anim_value * window_width);
    }
    if (tasklist_buttons[0].x >= window_rect.right) { // taskbar on right
      x_offset = (int)(anim_value * window_width);
    }
    if (tasklist_buttons[0].y <= window_rect.top) { // taskbar on top
      y_offset = (int)(-anim_value * window_height);
    }
    if (tasklist_buttons[0].y >= window_rect.bottom) { // taskbar on bottom
      y_offset = (int)(anim_value * window_height);
    }
  } else {
    x_offset = 0;
    y_offset = anim_value * window_height;
  }
  // Draw background
  winrt::com_ptr<ID2D1SolidColorBrush> brush;
  D2D1_COLOR_F brushColor = colors.light_mode ? D2D1::ColorF(1.0f, 1.0f, 1.0f, 0.8f) : D2D1::ColorF(0, 0, 0, 0.8f);
  winrt::check_hresult(d2d_dc->CreateSolidColorBrush(brushColor, brush.put()));
  D2D1_RECT_F background_rect = {};
  background_rect.bottom = (float)window_height;
  background_rect.right = (float)window_width;
  d2d_dc->SetTransform(D2D1::Matrix3x2F::Identity());
  d2d_dc->FillRectangle(background_rect, brush.get());
 
  // Set the animation - move the draw window according to annimation step
  auto popin = D2D1::Matrix3x2F::Translation(x_offset, y_offset);
  d2d_dc->SetTransform(popin);

  // Thumbnail logic:
  auto thumb_window = get_window_pos(active_window);
  bool minature_shown = active_window != nullptr && thumbnail != nullptr && thumb_window;
  if (minature_shown && thumb_window->right - thumb_window->left <= 0 || thumb_window->bottom - thumb_window->top <= 0)
    minature_shown = false;
  bool render_monitors = true;
  auto total_monitor_with_screen = total_monitor;
  if (thumb_window) {
    total_monitor_with_screen.rect.left = min(total_monitor_with_screen.rect.left, thumb_window->left + monitor_dx);
    total_monitor_with_screen.rect.top = min(total_monitor_with_screen.rect.top, thumb_window->top + monitor_dy);
    total_monitor_with_screen.rect.right = max(total_monitor_with_screen.rect.right, thumb_window->right + monitor_dx);
    total_monitor_with_screen.rect.bottom = max(total_monitor_with_screen.rect.bottom, thumb_window->bottom + monitor_dy);
  }
  // Only allow the new rect beeing slight bigger.
  if (total_monitor_with_screen.width() - total_monitor.width() > (thumb_window->right - thumb_window->left) / 2 ||
      total_monitor_with_screen.height() - total_monitor.height() > (thumb_window->bottom - thumb_window->top) / 2) {
    render_monitors = false;
  }
  auto rect_and_scale = use_overlay->get_thumbnail_rect_and_scale(0, 0, total_monitor_with_screen.width(), total_monitor_with_screen.height(), 1);
  if (minature_shown) {
    RECT thumbnail_pos;
    if (render_monitors) {
      thumbnail_pos.left = (thumb_window->left + monitor_dx) * rect_and_scale.scale + rect_and_scale.rect.left;
      thumbnail_pos.top = (thumb_window->top + monitor_dy) * rect_and_scale.scale + rect_and_scale.rect.top;
      thumbnail_pos.right = (thumb_window->right + monitor_dx) * rect_and_scale.scale + rect_and_scale.rect.left;
      thumbnail_pos.bottom = (thumb_window->bottom + monitor_dy) * rect_and_scale.scale + rect_and_scale.rect.top;
    } else {
      thumbnail_pos = use_overlay->get_thumbnail_rect_and_scale(0, 0, thumb_window->right - thumb_window->left, thumb_window->bottom - thumb_window->top, 1).rect;
    }
    // If the animation is done show the thumbnail
    //   we cannot animate the thumbnail, the animation lags behind
    if (anim_value == 0) {
      minature_shown = show_thumbnail(thumbnail_pos);
    }
  }
  // render the monitors
  if (render_monitors) {
    brushColor = D2D1::ColorF(colors.start_color_menu, minature_shown ? 1.0 : 0.3);
    brush = nullptr;
    winrt::check_hresult(d2d_dc->CreateSolidColorBrush(brushColor, brush.put()));
    for (auto& monitor : monitors) {
      D2D1_RECT_F monitor_rect;
      monitor_rect.left = (monitor.rect.left + monitor_dx) * rect_and_scale.scale + rect_and_scale.rect.left;
      monitor_rect.top = (monitor.rect.top + monitor_dy) * rect_and_scale.scale + rect_and_scale.rect.top;
      monitor_rect.right = (monitor.rect.right + monitor_dx) * rect_and_scale.scale + rect_and_scale.rect.left;
      monitor_rect.bottom = (monitor.rect.bottom + monitor_dy)  * rect_and_scale.scale + rect_and_scale.rect.top;
      d2d_dc->FillRectangle(monitor_rect, brush.get());
    }
  }
  // Finalize the overlay - dimm the buttons if no thumbnail is present and show "No active window"
  use_overlay->toggle_window_group(minature_shown);
  if (!minature_shown) {
    no_active.render(d2d_dc);
  }
  // Finally: render the overlay...
  use_overlay->render(d2d_dc);
  // ... and the arrows with numbers
  for (auto&& button : tasklist_buttons) {
    if ((unsigned)button.keynum - 1 >= arrows.size())
      continue;
    render_arrow(arrows[button.keynum - 1], button, window_rect, use_overlay->get_scale(), d2d_dc);
  }
}
