#include "pch.h"
#include "d2d_window.h"
#include "monitors.h"
#include "utils.h"
#include <d2d1helper.h>
#include <dwmapi.h>

#pragma comment(lib, "dxgi")
#pragma comment(lib, "d3d11")
#pragma comment(lib, "d2d1")
#pragma comment(lib, "dcomp")
#pragma comment(lib, "dwmapi")

extern "C" IMAGE_DOS_HEADER __ImageBase;



struct WINDOWCOMPOSITIONATTRIBDATA {
  DWORD Attrib;
  PVOID pvData;
  SIZE_T cbData;
};

struct ACCENT_POLICY {
  DWORD AccentState;
  DWORD AccentFlags;
  DWORD GradientColor;
  DWORD AnimationId;
};

typedef BOOL(WINAPI*pfnSetWindowCompositionAttribute)(HWND, WINDOWCOMPOSITIONATTRIBDATA*);

pfnSetWindowCompositionAttribute getSetWindowCompositionAttributeFunPtr() {
  auto user32 = LoadLibrary("user32.dll");
  auto rval =  reinterpret_cast<pfnSetWindowCompositionAttribute>(GetProcAddress(user32, "SetWindowCompositionAttribute"));
  FreeLibrary(user32);
  return rval;
}

void enable_acrylic_window(HWND hwnd) {
  auto SetWindowCompositionAttribute = getSetWindowCompositionAttributeFunPtr();
  ACCENT_POLICY accent = {};
  accent.AccentState = 3; // ACCENT_ENABLE_BLURBEHIND;
  accent.GradientColor = 0xffffff;
  accent.AccentFlags = 0;
  WINDOWCOMPOSITIONATTRIBDATA data;
  data.Attrib = 19; // WCA_ACCENT_POLICY
  data.pvData = &accent;
  data.cbData = sizeof(accent);
  SetWindowCompositionAttribute(hwnd, &data);
}


D2DWindow::D2DWindow() {
  static const char* class_name = "PToyD2DPopup";
  auto primary_screen = get_primary_monitor();
  WNDCLASS wc = {};
  wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
  wc.hInstance = reinterpret_cast<HINSTANCE>(&__ImageBase);
  wc.lpszClassName = class_name;
  wc.style = CS_HREDRAW | CS_VREDRAW;
  wc.lpfnWndProc = d2d_window_proc;
  RegisterClass(&wc);
  hwnd = CreateWindowEx(WS_EX_TOOLWINDOW | WS_EX_TOPMOST | WS_EX_NOREDIRECTIONBITMAP,
                        wc.lpszClassName,
                        "PToyD2DPopup",
                        WS_POPUP| WS_VISIBLE,
                        CW_USEDEFAULT, CW_USEDEFAULT,
                        CW_USEDEFAULT, CW_USEDEFAULT,
                        nullptr, nullptr, wc.hInstance, this);
  WINRT_VERIFY(hwnd);
  init();
  enable_acrylic_window(hwnd);
  MoveWindow(hwnd, primary_screen.left(), primary_screen.top(), primary_screen.width(), primary_screen.height(), TRUE);
}

/*
void D2DWindow::init() {
  auto primary_screen = get_primary_monitor();
  
  
}
*/
void D2DWindow::init() {
  // D2D1Factory is independent from the device, no need to recreate it if
  // we need to recreate the device.
  if (!d2d_factory) {
    D2D1_FACTORY_OPTIONS options = { D2D1_DEBUG_LEVEL_INFORMATION };
    winrt::check_hresult(D2D1CreateFactory(
      D2D1_FACTORY_TYPE_SINGLE_THREADED,
      __uuidof(d2d_factory),
      &options,
      d2d_factory.put_void()));
  }
  // For all other stuff - assing nullptr first to release the object, to reset
  // the com_ptr.
  d3d_device = nullptr;
  winrt::check_hresult(D3D11CreateDevice(
    nullptr,
    D3D_DRIVER_TYPE_HARDWARE,
    nullptr,
    D3D11_CREATE_DEVICE_BGRA_SUPPORT,
    nullptr,
    0,
    D3D11_SDK_VERSION,
    d3d_device.put(),
    nullptr,
    nullptr));
  // A bug in winrt - d3d_device.as(dxgi_device) does not compile. Roll out our own implementation:
  dxgi_device = nullptr;
  winrt::check_hresult(d3d_device->QueryInterface(
    __uuidof(dxgi_device),
    dxgi_device.put_void()));
  dxgi_factory = nullptr;
  winrt::check_hresult(CreateDXGIFactory2(
    DXGI_CREATE_FACTORY_DEBUG,
    __uuidof(dxgi_factory),
    dxgi_factory.put_void()));
  d2d_device = nullptr;
  winrt::check_hresult(d2d_factory->CreateDevice(dxgi_device.get(), d2d_device.put()));
  d2d_dc = nullptr;
  winrt::check_hresult(d2d_device->CreateDeviceContext(
    D2D1_DEVICE_CONTEXT_OPTIONS_NONE,
    d2d_dc.put()));
  svg_strem = nullptr;
  winrt::check_hresult(SHCreateStreamOnFileEx(
    L"overlay.svg",
    STGM_READ,
    FILE_ATTRIBUTE_NORMAL,
    FALSE,
    nullptr,
    svg_strem.put()));
  svg_document = nullptr;
  winrt::check_hresult(d2d_dc->CreateSvgDocument(
    svg_strem.get(),
    D2D1::SizeF(1,1),
    svg_document.put()));
}

void D2DWindow::resize() {
  auto window_rect = *get_window_pos(hwnd);
  auto width = window_rect.right - window_rect.left;
  auto height = window_rect.bottom - window_rect.top;
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
  winrt::check_hresult(dxgi_factory->CreateSwapChainForComposition(
    dxgi_device.get(),
    &sc_description,
    nullptr,
    dxgi_swap_chain.put()));
  composition_device = nullptr;
  winrt::check_hresult(DCompositionCreateDevice(
    dxgi_device.get(),
    __uuidof(composition_device),
    composition_device.put_void()));

  target = nullptr;
  winrt::check_hresult(composition_device->CreateTargetForHwnd(hwnd, true, target.put()));
  visual = nullptr;
  winrt::check_hresult(composition_device->CreateVisual(visual.put()));
  winrt::check_hresult(visual->SetContent(dxgi_swap_chain.get()));
  winrt::check_hresult(target->SetRoot(visual.get()));
  
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
}

HTHUMBNAIL tid = nullptr;
void D2DWindow::render() {
 
    if (!d2d_dc || !d2d_bitmap)
      return;

    auto active = FindWindow("notepad", NULL); //GetForegroundWindow();
    if (active && tid == nullptr) {
      winrt::check_hresult(DwmRegisterThumbnail(hwnd, active, &tid));
      RECT dest = { 0,0,100,150 };
      DWM_THUMBNAIL_PROPERTIES dskThumbProps;
      dskThumbProps.dwFlags = DWM_TNP_SOURCECLIENTAREAONLY | DWM_TNP_VISIBLE | DWM_TNP_OPACITY | DWM_TNP_RECTDESTINATION;
      dskThumbProps.fSourceClientAreaOnly = FALSE;
      dskThumbProps.fVisible = TRUE;
      dskThumbProps.opacity = (255 * 70) / 100;
      dskThumbProps.rcDestination = dest;

      // Display the thumbnail
      winrt::check_hresult(DwmUpdateThumbnailProperties(tid, &dskThumbProps));
    }

    d2d_dc->BeginDraw();
    d2d_dc->Clear();

    // Draw background
    d2d_dc->SetTransform(D2D1::Matrix3x2F::Identity());
    winrt::com_ptr<ID2D1SolidColorBrush> brush;
    D2D1_COLOR_F const brushColor = D2D1::ColorF(1.0f, 1.0f, 1.0f, 0.8f);
    winrt::check_hresult(d2d_dc->CreateSolidColorBrush(brushColor, brush.put()));
    D2D1_RECT_F rect;
    rect.left = 0; rect.top = 0;
    rect.bottom = 2160;
    rect.right = 3840;
    d2d_dc->FillRectangle(rect, brush.get());
    // Draw SVG
    D2D1_MATRIX_3X2_F transform = D2D1::Matrix3x2F::Identity();
    transform = transform * D2D1::Matrix3x2F::Translation((3840 - 1258) / 2, (2160 - 554) / 2);
    transform = transform * D2D1::Matrix3x2F::Scale(2.5, 2.5, D2D1::Point2F(1920, 1080));
    d2d_dc->SetTransform(transform);
    d2d_dc->DrawSvgDocument(svg_document.get());

    winrt::check_hresult(d2d_dc->EndDraw());

    winrt::check_hresult(dxgi_swap_chain->Present(1, 0));
    winrt::check_hresult(composition_device->Commit());
}

D2DWindow::~D2DWindow() {
  DestroyWindow(hwnd);
}
 
D2DWindow* D2DWindow::this_from_hwnd(HWND window) {
  return reinterpret_cast<D2DWindow*>(GetWindowLongPtr(window, GWLP_USERDATA));
}

LRESULT __stdcall D2DWindow::d2d_window_proc(HWND window, UINT message, WPARAM wparam, LPARAM lparam) {
  switch (message) {
  case WM_NCCREATE: {
    auto create_struct = reinterpret_cast<CREATESTRUCT*>(lparam);
    SetWindowLongPtr(window, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create_struct->lpCreateParams));
    return TRUE;
  }
  case WM_SIZE:
    this_from_hwnd(window)->resize();
  case WM_PAINT:
    this_from_hwnd(window)->render();
    return 0;
  case WM_DESTROY:
    PostQuitMessage(0);
    return 0;
  default:
    return DefWindowProc(window, message, wparam, lparam);
  }
}
