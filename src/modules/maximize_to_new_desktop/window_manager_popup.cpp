#include "pch.h"
#include "window_manager_popup.h"
#include <common/monitors.h>
#include "virtual_desktops.h"
#include "mouse_track_events.h"
#include <common/windows_colors.h>
#include <winrt/Windows.UI.ViewManagement.h>

extern "C" IMAGE_DOS_HEADER __ImageBase;

const LPWSTR maximize_tooltip_message = (LPWSTR)L"Maximize to new desktop";
const LPWSTR restore_tooltip_message = (LPWSTR)L"Return to primary desktop";

D2DWindowManagerPopup::D2DWindowManagerPopup() {
  static const TCHAR* class_name = L"PToyD2DWindowManagerPopup";
  WNDCLASS wc = {};
  wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
  wc.hInstance = reinterpret_cast<HINSTANCE>(&__ImageBase);
  wc.lpszClassName = class_name;
  wc.style = CS_HREDRAW | CS_VREDRAW;
  wc.lpfnWndProc = d2d_window_proc;
  RegisterClass(&wc);
  hwnd = CreateWindowEx(WS_EX_TOOLWINDOW | WS_EX_TOPMOST | WS_EX_LAYERED | WS_EX_NOACTIVATE,
                         wc.lpszClassName,
                         L"PToyD2DWindowManagerPopup",
                         WS_POPUP | WS_VISIBLE,
                         CW_USEDEFAULT, CW_USEDEFAULT,
                         CW_USEDEFAULT, CW_USEDEFAULT,
                         nullptr, nullptr, wc.hInstance, this);
  WINRT_VERIFY(hwnd);
  SetLayeredWindowAttributes(hwnd, 0, (int)(255), LWA_ALPHA);
  ShowWindow(hwnd,SW_HIDE);
  init();
}

D2DWindowManagerPopup::~D2DWindowManagerPopup() {
  hide();
  DestroyWindow(hwnd);
}

void D2DWindowManagerPopup::init()
{
  std::unique_lock lock(mutex);
  target_window = NULL;
  // D2D1Factory is independent from the device, no need to recreate it if
  // we need to recreate the device.
  if (!d2d_factory) {
    D2D1_FACTORY_OPTIONS options = { D2D1_DEBUG_LEVEL_INFORMATION };
    winrt::check_hresult(D2D1CreateFactory(D2D1_FACTORY_TYPE_MULTI_THREADED,
                                           __uuidof(d2d_factory),
                                           &options,
                                           d2d_factory.put_void()));
  }
  // For all other stuff - assing nullptr first to release the object, to reset
  // the com_ptr.
  d3d_device = nullptr;
  winrt::check_hresult(D3D11CreateDevice(nullptr,
                                         D3D_DRIVER_TYPE_HARDWARE,
                                         nullptr,
                                         D3D11_CREATE_DEVICE_BGRA_SUPPORT,
                                         nullptr,
                                         0,
                                         D3D11_SDK_VERSION,
                                         d3d_device.put(),
                                         nullptr,
                                         nullptr));
  dxgi_device = nullptr;
  winrt::check_hresult(d3d_device->QueryInterface(__uuidof(dxgi_device), dxgi_device.put_void()));
  dxgi_factory = nullptr;
  winrt::check_hresult(CreateDXGIFactory2(0, __uuidof(dxgi_factory), dxgi_factory.put_void()));
  d2d_device = nullptr;
  winrt::check_hresult(d2d_factory->CreateDevice(dxgi_device.get(), d2d_device.put()));
  d2d_dc = nullptr;
  winrt::check_hresult(d2d_device->CreateDeviceContext(D2D1_DEVICE_CONTEXT_OPTIONS_NONE, d2d_dc.put()));

  maximize_to_new_desktop.load(L"svgs\\maximize_to_new_desktop.svg", d2d_dc.get());
  restore_to_primary_desktop.load(L"svgs\\restore_to_primary_desktop.svg", d2d_dc.get());
  current_icon = &maximize_to_new_desktop;
}

D2DWindowManagerPopup* D2DWindowManagerPopup::this_from_hwnd(HWND window) {
  return reinterpret_cast<D2DWindowManagerPopup*>(GetWindowLongPtr(window, GWLP_USERDATA));
}

void D2DWindowManagerPopup::create_tooltip(HWND window) {
  //Create a tooltip.
  hwnd_tooltip = CreateWindowEx(NULL, TOOLTIPS_CLASS, NULL,
                                WS_POPUP | TTS_ALWAYSTIP | TTS_NOPREFIX ,
                                CW_USEDEFAULT, CW_USEDEFAULT,
                                CW_USEDEFAULT, CW_USEDEFAULT,
                                window, NULL,
                                reinterpret_cast<HINSTANCE>(&__ImageBase), NULL);

  SetWindowPos(hwnd_tooltip, HWND_TOPMOST,0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
  // Associate the tooltip with the tool.
  TOOLINFO tool_info = { 0 };
  tool_info.cbSize = TTTOOLINFO_V1_SIZE; // sizeof(TOOLINFO) doesn't work...
  tool_info.hwnd = window;
  tool_info.uFlags = TTF_SUBCLASS;
  tool_info.lpszText = maximize_tooltip_message;
  tool_info.hinst = reinterpret_cast<HINSTANCE>(&__ImageBase);
  GetClientRect(hwnd, &tool_info.rect);
  SendMessage(hwnd_tooltip, TTM_ADDTOOL, 0, (LPARAM)&tool_info);
}

HWND D2DWindowManagerPopup::get_hwnd() {
  return hwnd;
}

void D2DWindowManagerPopup::show(HWND targetWindow, RECT area) {
  int current_desktop_index = get_desktop_index_for_window(targetWindow);
  if (current_desktop_index < 0) {
    // Couldn't find the desktop.
    return;
  }
  target_window = targetWindow;
  current_icon = (current_desktop_index == 0 ? &maximize_to_new_desktop : &restore_to_primary_desktop);
  target_window_on_primary_desktop = (current_desktop_index == 0 ? TRUE : FALSE);

  TOOLINFO tool_info = { 0 };
  tool_info.cbSize = TTTOOLINFO_V1_SIZE;
  tool_info.hwnd = hwnd;
  // Get the current tooltip definition.
  if (SendMessage(hwnd_tooltip, TTM_GETTOOLINFO, 0, (LPARAM)&tool_info)) {
    // Change the tooltip text.
    tool_info.lpszText = target_window_on_primary_desktop ? maximize_tooltip_message : restore_tooltip_message;
    SendMessage(hwnd_tooltip, TTM_UPDATETIPTEXT, 0, (LPARAM)&tool_info);
  }

  SetWindowPos(hwnd, HWND_TOPMOST, area.left, area.top, area.right - area.left, area.bottom - area.top, 0);
  ShowWindow(hwnd, SW_SHOWNOACTIVATE);
}

void D2DWindowManagerPopup::hide() {
  ShowWindow(hwnd, SW_HIDE);
}

void D2DWindowManagerPopup::resize() {
  std::unique_lock lock(mutex);
  auto window_rect = *get_window_pos(hwnd);
  auto width = window_rect.right - window_rect.left;
  auto height = window_rect.bottom - window_rect.top;
  hwnd_rect.left = (float)0;
  hwnd_rect.top = (float)0;
  hwnd_rect.bottom = (float)height;
  hwnd_rect.right = (float)width;
  if (width == 0 || height == 0)
    return;
  DXGI_SWAP_CHAIN_DESC1 sc_description = {};
  sc_description.Format = DXGI_FORMAT_B8G8R8A8_UNORM;
  sc_description.BufferUsage = DXGI_USAGE_RENDER_TARGET_OUTPUT;
  sc_description.SwapEffect = DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL;
  sc_description.BufferCount = 2;
  sc_description.SampleDesc.Count = 1;
  sc_description.AlphaMode = DXGI_ALPHA_MODE_PREMULTIPLIED;
  sc_description.Width = width;
  sc_description.Height = height;
  dxgi_swap_chain = nullptr;
  winrt::check_hresult(dxgi_factory->CreateSwapChainForComposition(dxgi_device.get(),
                                                                   &sc_description,
                                                                   nullptr,
                                                                   dxgi_swap_chain.put()));
  composition_device = nullptr;
  winrt::check_hresult(DCompositionCreateDevice(dxgi_device.get(),
                                                __uuidof(composition_device),
                                                composition_device.put_void()));
  composition_target = nullptr;
  winrt::check_hresult(composition_device->CreateTargetForHwnd(hwnd, true, composition_target.put()));
  composition_visual = nullptr;
  winrt::check_hresult(composition_device->CreateVisual(composition_visual.put()));
  winrt::check_hresult(composition_visual->SetContent(dxgi_swap_chain.get()));
  winrt::check_hresult(composition_target->SetRoot(composition_visual.get()));

  dxgi_surface = nullptr;
  winrt::check_hresult(dxgi_swap_chain->GetBuffer(0, __uuidof(dxgi_surface), dxgi_surface.put_void()));
  D2D1_BITMAP_PROPERTIES1 properties = {};
  properties.pixelFormat.alphaMode = D2D1_ALPHA_MODE_PREMULTIPLIED;
  properties.pixelFormat.format = DXGI_FORMAT_B8G8R8A8_UNORM;
  properties.bitmapOptions = D2D1_BITMAP_OPTIONS_TARGET | D2D1_BITMAP_OPTIONS_CANNOT_DRAW;
  d2d_bitmap = nullptr;
  winrt::check_hresult(d2d_dc->CreateBitmapFromDxgiSurface(dxgi_surface.get(),
                                                           properties,
                                                           d2d_bitmap.put()));
  d2d_dc->SetTarget(d2d_bitmap.get());

  TOOLINFO tool_info = { 0 };
  tool_info.cbSize = TTTOOLINFO_V1_SIZE;
  tool_info.hwnd = hwnd;
  // Get the current tooltip definition.
  if(SendMessage(hwnd_tooltip, TTM_GETTOOLINFO, 0, (LPARAM)&tool_info)) {
    // Resize the tooltip rect.
    GetClientRect(hwnd, &tool_info.rect);
    SendMessage(hwnd_tooltip, TTM_NEWTOOLRECT, 0, (LPARAM)&tool_info);
  }
}

void D2DWindowManagerPopup::render() {
  std::unique_lock lock(mutex);
  if (!d2d_bitmap)
    return;
  d2d_dc->BeginDraw();
  d2d_dc->Clear();
  // Draw background
  winrt::com_ptr<ID2D1SolidColorBrush> brush;
  DWORD back_color;
  if(is_pressed) {
    back_color = WindowsColors::rgb_color(WindowsColors::get_accent_dark_1_color());
  } else if(should_highlight) {
    back_color = WindowsColors::rgb_color(WindowsColors::get_accent_light_1_color());
  } else {
    back_color = WindowsColors::rgb_color(WindowsColors::get_accent_color());
  }
  D2D1_COLOR_F brushColor = D2D1::ColorF(back_color, 1.0f);
  winrt::check_hresult(d2d_dc->CreateSolidColorBrush(brushColor, brush.put()));
  d2d_dc->FillRectangle(hwnd_rect, brush.get());
  // Draw SVG
  int width = (int)(hwnd_rect.right - hwnd_rect.left);
  int height = (int)(hwnd_rect.bottom - hwnd_rect.top);
  // Apply padding, 96 DPI screens is 4px (see on_mouse_in() in dllmain.cpp)
  int padding_x = (int)(width * 0.07f);
  int padding_y = (int)(height * 0.15f);
  current_icon->resize(padding_x,
                       padding_y,
                       width - (2 * padding_x),
                       height - (2 * padding_y),
                       1.0f);
  DWORD icon_color;
  icon_color = WindowsColors::rgb_color(WindowsColors::get_highlight_text_color());
  D2D1_COLOR_F iconBrush = D2D1::ColorF(icon_color, 1.0f);
  current_icon->find_element(L"icon-visual")->SetAttributeValue(L"fill", iconBrush);
  current_icon->render(d2d_dc.get());
  winrt::check_hresult(d2d_dc->EndDraw());
  winrt::check_hresult(dxgi_swap_chain->Present(1, 0));
  winrt::check_hresult(composition_device->Commit());
}


LRESULT __stdcall D2DWindowManagerPopup::d2d_window_proc(HWND window, UINT message, WPARAM wparam, LPARAM lparam) {
  D2DWindowManagerPopup* _this;
  switch (message) {
  case WM_NCCREATE: {
    auto create_struct = reinterpret_cast<CREATESTRUCT*>(lparam);
    SetWindowLongPtr(window, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create_struct->lpCreateParams));
    return TRUE;
  }
  case WM_CREATE: {
    this_from_hwnd(window)->create_tooltip(window);
    return DefWindowProc(window, message, wparam, lparam);
  }
  case WM_SIZE:
    this_from_hwnd(window)->resize();
  case WM_PAINT:
    this_from_hwnd(window)->render();
    return DefWindowProc(window, message, wparam, lparam);
  case WM_LBUTTONDOWN:
    _this = this_from_hwnd(window);
    _this->is_pressed = true;
    InvalidateRect(window, NULL, TRUE);
    return DefWindowProc(window, message, wparam, lparam);
  case WM_LBUTTONUP:
    _this = this_from_hwnd(window);
    _this->hide();
    if (_this->target_window_on_primary_desktop) {
      move_window_to_new_desktop(_this->target_window);
    } else {
      move_window_to_primary_desktop(_this->target_window, _this->delete_after_restore);
    }
    return 0;
  case WM_MOUSEMOVE:
    // Enter the window
    _this = this_from_hwnd(window);
    _this->is_pressed = wparam & MK_LBUTTON;
    _this->should_highlight = true;
    _this->mouse_track.on_moude_move(window, TME_LEAVE); // Start tracking to see when we leave.
    InvalidateRect(window, NULL, TRUE);
    return DefWindowProc(window, message, wparam, lparam);
  case WM_MOUSELEAVE:
    // Leave the window
    _this = this_from_hwnd(window);
    _this->should_highlight = false;
    _this->is_pressed = false;
    _this->mouse_track.reset(window);
    InvalidateRect(window, NULL, TRUE);
    return DefWindowProc(window, message, wparam, lparam);
  default:
    return DefWindowProc(window, message, wparam, lparam);
  }
}

void D2DWindowManagerPopup::set_delete_after_restore(bool value)   {
  this->delete_after_restore = value;
}

bool D2DWindowManagerPopup::get_delete_after_restore() {
  return this->delete_after_restore;
}
