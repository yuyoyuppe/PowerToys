#include "pch.h"
#include "d2d_window.h"
#include "monitors.h"

#pragma comment(lib, "dxgi")
#pragma comment(lib, "d3d11")
#pragma comment(lib, "d2d1")
#pragma comment(lib, "dcomp")

extern "C" IMAGE_DOS_HEADER __ImageBase;

typedef enum _WINDOWCOMPOSITIONATTRIB
{
  WCA_UNDEFINED = 0,
  WCA_NCRENDERING_ENABLED = 1,
  WCA_NCRENDERING_POLICY = 2,
  WCA_TRANSITIONS_FORCEDISABLED = 3,
  WCA_ALLOW_NCPAINT = 4,
  WCA_CAPTION_BUTTON_BOUNDS = 5,
  WCA_NONCLIENT_RTL_LAYOUT = 6,
  WCA_FORCE_ICONIC_REPRESENTATION = 7,
  WCA_EXTENDED_FRAME_BOUNDS = 8,
  WCA_HAS_ICONIC_BITMAP = 9,
  WCA_THEME_ATTRIBUTES = 10,
  WCA_NCRENDERING_EXILED = 11,
  WCA_NCADORNMENTINFO = 12,
  WCA_EXCLUDED_FROM_LIVEPREVIEW = 13,
  WCA_VIDEO_OVERLAY_ACTIVE = 14,
  WCA_FORCE_ACTIVEWINDOW_APPEARANCE = 15,
  WCA_DISALLOW_PEEK = 16,
  WCA_CLOAK = 17,
  WCA_CLOAKED = 18,
  WCA_ACCENT_POLICY = 19,
  WCA_FREEZE_REPRESENTATION = 20,
  WCA_EVER_UNCLOAKED = 21,
  WCA_VISUAL_OWNER = 22,
  WCA_LAST = 23
} WINDOWCOMPOSITIONATTRIB;

typedef struct _WINDOWCOMPOSITIONATTRIBDATA
{
  WINDOWCOMPOSITIONATTRIB Attrib;
  PVOID pvData;
  SIZE_T cbData;
} WINDOWCOMPOSITIONATTRIBDATA;

typedef enum _ACCENT_STATE
{
  ACCENT_DISABLED = 0,
  ACCENT_ENABLE_GRADIENT = 1,
  ACCENT_ENABLE_TRANSPARENTGRADIENT = 2,
  ACCENT_ENABLE_BLURBEHIND = 3,
  ACCENT_INVALID_STATE = 4
} ACCENT_STATE;

typedef struct _ACCENT_POLICY
{
  ACCENT_STATE AccentState;
  DWORD AccentFlags;
  DWORD GradientColor;
  DWORD AnimationId;
} ACCENT_POLICY;


//WINUSERAPI BOOL WINAPI SetWindowCompositionAttribute(_In_ HWND hWnd, _Inout_ WINDOWCOMPOSITIONATTRIBDATA* pAttrData);
typedef BOOL(WINAPI*pfnSetWindowCompositionAttribute)(HWND, WINDOWCOMPOSITIONATTRIBDATA*);
pfnSetWindowCompositionAttribute SetWindowCompositionAttribute;

void undocumented_acrylic(HWND hwnd) {
  if (!SetWindowCompositionAttribute) {
    auto user32 = LoadLibrary("user32.dll");
    SetWindowCompositionAttribute = (pfnSetWindowCompositionAttribute)GetProcAddress(user32, "SetWindowCompositionAttribute");
  }
  ACCENT_POLICY accent = {};
  accent.AccentState = ACCENT_ENABLE_BLURBEHIND;
  WINDOWCOMPOSITIONATTRIBDATA data;
  data.Attrib = WCA_ACCENT_POLICY;
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
  undocumented_acrylic(hwnd);
  

  MoveWindow(hwnd, primary_screen.left(), primary_screen.top(), primary_screen.width(), primary_screen.height(), TRUE);
}

void D2DWindow::init() {
  auto primary_screen = get_primary_monitor();
  winrt::check_hresult(D3D11CreateDevice(nullptr,
                                         D3D_DRIVER_TYPE_HARDWARE,
                                         nullptr,
                                         D3D11_CREATE_DEVICE_BGRA_SUPPORT,
                                         nullptr, 0,
                                         D3D11_SDK_VERSION,
                                         d3d_device.put(),
                                         nullptr,
                                         nullptr));
  // A bug in winrt - d3d_device.as(dxgi_device) does not compile. Roll out our own implementation:
  winrt::check_hresult(d3d_device->QueryInterface(__uuidof(dxgi_device), dxgi_device.put_void()));

  winrt::check_hresult(CreateDXGIFactory2(DXGI_CREATE_FACTORY_DEBUG,
                                          __uuidof(dxgi_factory),
                                          dxgi_factory.put_void()));
  DXGI_SWAP_CHAIN_DESC1 sc_description = {};
  sc_description.Format = DXGI_FORMAT_B8G8R8A8_UNORM;
  sc_description.BufferUsage = DXGI_USAGE_RENDER_TARGET_OUTPUT;
  sc_description.SwapEffect = DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL;
  sc_description.BufferCount = 2;
  sc_description.SampleDesc.Count = 1;
  sc_description.AlphaMode = DXGI_ALPHA_MODE_PREMULTIPLIED;
  sc_description.Width = primary_screen.width();
  sc_description.Height = primary_screen.height();
  winrt::check_hresult(dxgi_factory->CreateSwapChainForComposition(dxgi_device.get(),
                                                                   &sc_description,
                                                                   nullptr,
                                                                   dxgi_swap_chain.put()));
  winrt::check_hresult(DCompositionCreateDevice(dxgi_device.get(),
                                                __uuidof(composition_device),
                                                composition_device.put_void()));
  
  winrt::check_hresult(composition_device->CreateTargetForHwnd(hwnd, true, target.put()));
  winrt::check_hresult(composition_device->CreateVisual(visual.put()));
  winrt::check_hresult(visual->SetContent(dxgi_swap_chain.get()));
  winrt::check_hresult(target->SetRoot(visual.get()));
}

void D2DWindow::render() {
 
    if (!composition_device)
      return;
    winrt::com_ptr<ID2D1Factory2> d2d_factory;
    D2D1_FACTORY_OPTIONS options = { D2D1_DEBUG_LEVEL_INFORMATION };
    winrt::check_hresult(D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED,
      options,
      d2d_factory.put()));

    winrt::com_ptr<ID2D1Device1> d2d_device;
    winrt::check_hresult(d2d_factory->CreateDevice(dxgi_device.get(), d2d_device.put()));

    winrt::com_ptr<ID2D1DeviceContext> d2d_dc;
    winrt::check_hresult(d2d_device->CreateDeviceContext(D2D1_DEVICE_CONTEXT_OPTIONS_NONE,
      d2d_dc.put()));
    winrt::com_ptr<IDXGISurface2> dxgi_surface;
    winrt::check_hresult(dxgi_swap_chain->GetBuffer(0, __uuidof(dxgi_surface), dxgi_surface.put_void()));

    D2D1_BITMAP_PROPERTIES1 properties = {};
    properties.pixelFormat.alphaMode = D2D1_ALPHA_MODE_PREMULTIPLIED;
    properties.pixelFormat.format = DXGI_FORMAT_B8G8R8A8_UNORM;
    properties.bitmapOptions = D2D1_BITMAP_OPTIONS_TARGET | D2D1_BITMAP_OPTIONS_CANNOT_DRAW;
    winrt::com_ptr<ID2D1Bitmap1> d2d_bitmap;
    winrt::check_hresult(d2d_dc->CreateBitmapFromDxgiSurface(dxgi_surface.get(),
      properties,
      d2d_bitmap.put()));

    d2d_dc->SetTarget(d2d_bitmap.get());
    d2d_dc->BeginDraw();
    d2d_dc->Clear();
    winrt::com_ptr<ID2D1SolidColorBrush> brush;
    D2D1_COLOR_F const brushColor = D2D1::ColorF(0.18f, 0.55f, 0.34f, 1.0f);
    winrt::check_hresult(d2d_dc->CreateSolidColorBrush(brushColor,
      brush.put()));
    D2D1_POINT_2F const ellipseCenter = D2D1::Point2F(150.0f, 150.0f);
    D2D1_ELLIPSE const ellipse = D2D1::Ellipse(ellipseCenter, 100.0f, 100.0f);
    d2d_dc->FillEllipse(ellipse, brush.get());
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
  case WM_PAINT:
  case WM_SIZE:
    this_from_hwnd(window)->render();
    return 0;
  case WM_DESTROY:
    PostQuitMessage(0);
    return 0;
  default:
    return DefWindowProc(window, message, wparam, lparam);
  }
}
