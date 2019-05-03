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
#include "d2d_window.h"
#include "mouse_track_events.h"

class D2DWindowManagerPopup
{
public:
  D2DWindowManagerPopup();
  void show(HWND targetWindow, RECT area);
  void hide();
  HWND get_hwnd();
  ~D2DWindowManagerPopup();
private:
  void init();
  static LRESULT __stdcall d2d_window_proc(HWND window, UINT message, WPARAM wparam, LPARAM lparam);
  static D2DWindowManagerPopup* this_from_hwnd(HWND window);
  void resize();
  void render();
  void create_tooltip(HWND window);

  HWND target_window;
  BOOL target_window_on_primary_desktop;

  HWND hwnd;
  HWND hwnd_tooltip;
  std::mutex mutex;
  D2D1_RECT_F hwnd_rect;

  D2DSVG maximize_to_new_desktop;
  D2DSVG restore_to_primary_desktop;
  D2DSVG* current_icon;

  MouseTrackEvents mouse_track;
  bool should_highlight = false;

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
