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

class Animation {
public:
  Animation(double duration) :
    duration(duration),
    start(std::chrono::high_resolution_clock::now()) { }
  void reset() {
    start = std::chrono::high_resolution_clock::now();
  }
  double value() const {
    auto anim_duration = std::chrono::high_resolution_clock::now() - start;
    double seconds = std::chrono::duration<double>(anim_duration).count() / duration;
    if (seconds > 1)
      seconds = 1;
    return 1 - pow(2, -8 * seconds);
  }
private:
  std::chrono::high_resolution_clock::time_point start;
  double duration;
};

class D2DWindow
{
public:
  D2DWindow();
  void show(int x, int y, int width, int height);
  void hide();
  ~D2DWindow();
protected:
  static LRESULT __stdcall d2d_window_proc(HWND window, UINT message, WPARAM wparam, LPARAM lparam);
  static D2DWindow* this_from_hwnd(HWND window);
  virtual std::unique_lock<std::mutex> init();
  virtual std::unique_lock<std::mutex> resize();
  virtual void render(ID2D1DeviceContext5* d2d_dc);
  void render_impl();

  std::mutex mutex;
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
