#pragma once
#include <winrt/base.h>
#include <Windows.h>
#include <dxgi1_3.h>
#include <d3d11_2.h>
#include <d2d1_3.h>
#include <d2d1_3helper.h>
#include <d2d1helper.h>
#include <dcomp.h>
#include <dwmapi.h>
#include <string>
#include "tasklist_positions.h"
#include "d2d_svg.h"
#include <winrt/Windows.UI.ViewManagement.h>

/*
  Usage:
    When creating animation contstructor takes one parameter - how long
    should the animation take in seconds.

    Call reset() when starting animation.

    When redering, call value() to get value from 0 to 1 - depending on animation
    progress.
*/
class Animation {
public:
  enum AnimFunctions {
    LINEAR = 0,
    EASE_OUT_EXPO
  };

  Animation(double duration = 1, double start = 0, double stop = 1);
  void reset();
  void reset(double duration);
  void reset(double duration, double start, double stop);
  double value(AnimFunctions apply_function) const;
  bool done() const;
private:
  double apply_animation_function(double t, AnimFunctions apply_function) const;
  std::chrono::high_resolution_clock::time_point start;
  double start_value, end_value, duration;
};

struct WindowsColors {
  using Color = winrt::Windows::UI::Color;
  
  static DWORD packed_color_from_ui_color(winrt::Windows::UI::Color color) {
    return ((DWORD)color.R << 16) | ((DWORD)color.G << 8) | ((DWORD)color.B);
  }
  static winrt::Windows::UI::Color get_button_face_color() {
    winrt::Windows::UI::ViewManagement::UISettings uiSettings;
    return uiSettings.UIElementColor(winrt::Windows::UI::ViewManagement::UIElementType::ButtonFace);
  }
  static winrt::Windows::UI::Color get_button_text_color() {
    winrt::Windows::UI::ViewManagement::UISettings uiSettings;
    return uiSettings.UIElementColor(winrt::Windows::UI::ViewManagement::UIElementType::ButtonText);
  }
  static winrt::Windows::UI::Color get_highlight_color() {
    winrt::Windows::UI::ViewManagement::UISettings uiSettings;
    return uiSettings.UIElementColor(winrt::Windows::UI::ViewManagement::UIElementType::Highlight);
  }
  static winrt::Windows::UI::Color get_hotlight_color() {
    winrt::Windows::UI::ViewManagement::UISettings uiSettings;
    return uiSettings.UIElementColor(winrt::Windows::UI::ViewManagement::UIElementType::Hotlight);
  }
  static winrt::Windows::UI::Color get_highlight_text_color() {
    winrt::Windows::UI::ViewManagement::UISettings uiSettings;
    return uiSettings.UIElementColor(winrt::Windows::UI::ViewManagement::UIElementType::HighlightText);
  }
  static winrt::Windows::UI::Color get_accent_light_1_color() {
    winrt::Windows::UI::ViewManagement::UISettings uiSettings;
    return uiSettings.GetColorValue(winrt::Windows::UI::ViewManagement::UIColorType::AccentLight1);
  }
  static winrt::Windows::UI::Color get_accent_light_2_color() {
    winrt::Windows::UI::ViewManagement::UISettings uiSettings;
    return uiSettings.GetColorValue(winrt::Windows::UI::ViewManagement::UIColorType::AccentLight2);
  }
  static winrt::Windows::UI::Color get_accent_color() {
    winrt::Windows::UI::ViewManagement::UISettings uiSettings;
    return uiSettings.GetColorValue(winrt::Windows::UI::ViewManagement::UIColorType::Accent);
  }
  static winrt::Windows::UI::Color get_background_color() {
    winrt::Windows::UI::ViewManagement::UISettings uiSettings;
    return uiSettings.GetColorValue(winrt::Windows::UI::ViewManagement::UIColorType::Background);
  }
  static DWORD unpack_color(DWORD color) {
    // registry keeps the colors in ABGR format, we want RGB
    auto r = (color & 0xFF);
    auto g = (color & 0xFF00) >> 8;
    auto b = (color & 0xFF0000) >> 16;
    return (r << 16) | (g << 8) | b;
  }
  // Update colors - returns true if the values where changed
  bool update() {
    DWORD data_size = sizeof(DWORD);
    DWORD new_accent_color_menu = 0;
    DWORD new_start_color_menu = 0;
    DWORD new_desktop_fill_color = 0;
    bool new_light_mode = true;
    new_accent_color_menu = packed_color_from_ui_color(get_accent_color());
    new_start_color_menu = new_accent_color_menu;
    DWORD light_reg_val = packed_color_from_ui_color(get_background_color());
    new_desktop_fill_color = unpack_color(GetSysColor(COLOR_DESKTOP));

    new_light_mode = light_reg_val != 0; //Dark mode will have black as the background color.

    bool changed = new_accent_color_menu != accent_color_menu  ||
                   new_start_color_menu != start_color_menu ||
                   new_light_mode != light_mode ||
                   new_desktop_fill_color != desktop_fill_color;
    accent_color_menu = new_accent_color_menu;
    start_color_menu = new_start_color_menu;
    light_mode = new_light_mode;
    desktop_fill_color = new_desktop_fill_color;
    
    return changed;
  }
  DWORD accent_color_menu = 0, start_color_menu = 0, desktop_fill_color = 0;
  bool light_mode = true;
};

class D2DWindow
{
public:
  D2DWindow();
  void show(int x, int y, int width, int height);
  void hide();
  void initialize();
  ~D2DWindow();
protected:
  // Implement this:

  // Initialization - called when D2D device needs to be created.
  //   When called all D2DWindow members will be initialized, including d2d_dc
  virtual void init() = 0;
  // resize - when called, window_width and window_height will have current window size
  virtual void resize() = 0;
  // render - called on WM_PAIT, BeginPaint/EndPaint is handled by D2DWindow
  virtual void render(ID2D1DeviceContext5* d2d_dc) = 0;
  // on_show, on_hide - called when the window is about to be shown or about to be hidden
  virtual void on_show() = 0;
  virtual void on_hide() = 0;

  static LRESULT __stdcall d2d_window_proc(HWND window, UINT message, WPARAM wparam, LPARAM lparam);
  static D2DWindow* this_from_hwnd(HWND window);
  
  void base_init();
  void base_resize(int width, int height);
  void base_render();
  void render_empty();

  std::recursive_mutex mutex;
  bool initialized = false;
  HWND hwnd;
  long window_width, window_height;
  winrt::com_ptr<ID3D11Device> d3d_device;
  winrt::com_ptr<IDXGIDevice> dxgi_device;
  winrt::com_ptr<IDXGIFactory2> dxgi_factory;
  winrt::com_ptr<IDXGISwapChain1> dxgi_swap_chain;
  winrt::com_ptr<IDCompositionDevice> composition_device;
  winrt::com_ptr<IDCompositionTarget> composition_target;
  winrt::com_ptr<IDCompositionVisual> composition_visual;
  winrt::com_ptr<IDXGISurface2> dxgi_surface;
  winrt::com_ptr<ID2D1Bitmap1> d2d_bitmap;
  winrt::com_ptr<ID2D1Factory6> d2d_factory;
  winrt::com_ptr<ID2D1Device5> d2d_device;
  winrt::com_ptr<ID2D1DeviceContext5> d2d_dc;
};
