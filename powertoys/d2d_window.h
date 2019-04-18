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
  D2DSVG& resize(int x, int y, int width, int height, float fill, float max_scale = -1.0f);
  D2DSVG& render(ID2D1DeviceContext5* d2d_dc);
  float get_scale() const { return used_scale; }
  int width() const { return svg_width; }
  int height() const { return svg_height; }
  D2DSVG& toggle_element(const wchar_t* id, bool visible);
protected:
  float used_scale;
  winrt::com_ptr<ID2D1SvgDocument> svg;
  int svg_width, svg_height;
  D2D1::Matrix3x2F transform;
};

class D2DOverlaySVG : public D2DSVG {
public:
  D2DOverlaySVG& load(const std::wstring& filename, ID2D1DeviceContext5* d2d_dc);
  D2DOverlaySVG& resize(int x, int y, int width, int height, float fill, float max_scale = -1.0f);
  D2DOverlaySVG& find_thumbnail(const std::wstring& id);
  D2DOverlaySVG& find_window_group(const std::wstring& id);
  RECT get_thumbnail_rect(int window_cx, int window_cy, float scale);
  D2DOverlaySVG& toggle_window_group(bool active);
private:
  D2D1_POINT_2F thumbnail_top_left, thumbnail_bottom_right;
  RECT thumbnail_scaled_rect;
  winrt::com_ptr<ID2D1SvgElement> window_group;

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
  long window_width, window_height;
  HTHUMBNAIL thumbnail;

  D2DOverlaySVG landscape, portrait;
  D2DOverlaySVG* use_overlay;
  D2DSVG no_active;
  std::vector<D2DSVG> arrows;

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
