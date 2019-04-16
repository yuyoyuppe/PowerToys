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

class D2DWindow
{
public:
  D2DWindow();
  void show(HWND active_window);
  void hide();
  ~D2DWindow();
private:
  static LRESULT __stdcall d2d_window_proc(HWND window, UINT message, WPARAM wparam, LPARAM lparam);
  static D2DWindow* this_from_hwnd(HWND window);
  void init();
  void resize();
  bool show_thumbnail();
  void render();

  std::mutex mutex;
  HWND hwnd;
  D2D1_RECT_F hwnd_rect;
  D2D1_POINT_2F thumbnail_top_left, thumbnail_bottom_right;
  RECT thumbnail_scaled_rect;
  HTHUMBNAIL thumbnail;
  winrt::com_ptr<IStream> svg_strem;
  winrt::com_ptr<ID2D1SvgDocument> svg_document;
  winrt::com_ptr<ID2D1SvgElement> svg_window_group;
  int svg_width, svg_height;
  D2D1_MATRIX_3X2_F svg_rescale;
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
