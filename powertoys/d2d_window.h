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

class D2DSVG {
public:
  D2DSVG& load(const std::wstring& filename, ID2D1DeviceContext5* d2d_dc);
  D2DSVG& find_thumbnail(const std::wstring& id);
  D2DSVG& find_window_group(const std::wstring& id);
  D2DSVG& resize(int x, int y, int width, int height, float fill);
  RECT get_thumbnail_rect(int window_cx, int window_cy);
  D2DSVG& toggle_window_group(bool active);
  D2DSVG& render(ID2D1DeviceContext5* d2d_dc);
private:
  D2D1_POINT_2F thumbnail_top_left, thumbnail_bottom_right;
  RECT thumbnail_scaled_rect;
  winrt::com_ptr<ID2D1SvgDocument> svg;
  winrt::com_ptr<ID2D1SvgElement> window_group;
  int svg_width, svg_height;
  D2D1_MATRIX_3X2_F transform;
};

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
  HTHUMBNAIL thumbnail;

  D2DSVG landscape, portrait;
  D2DSVG* use_overlay;

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
