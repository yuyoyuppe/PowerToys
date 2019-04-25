#include "pch.h"
#include "d2d_window.h"
#include "monitors.h"
#include "tasklist_positions.h"

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
  static auto SetWindowCompositionAttribute = getSetWindowCompositionAttributeFunPtr();
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
}

void D2DWindow::show(HWND active_window) {
  tasklist.update();
  if (active_window) {
    // Ignore errors, if this fails we will just not show the thumbnail
    DwmRegisterThumbnail(hwnd, active_window, &thumbnail);
  }
  auto primary_screen = get_primary_monitor();
  SetWindowPos(hwnd, HWND_TOPMOST, primary_screen.left(), primary_screen.top(), primary_screen.width(), primary_screen.height(), 0);
  ShowWindow(hwnd, SW_SHOWNORMAL);
}

void D2DWindow::hide() {
  if (thumbnail) {
    DwmUnregisterThumbnail(thumbnail);
  }
  ShowWindow(hwnd, SW_HIDE);
}

D2DSVG& D2DSVG::load(const std::wstring& filename, ID2D1DeviceContext5* d2d_dc) {
  svg = nullptr;
  winrt::com_ptr<IStream> svg_stream;
  winrt::check_hresult(SHCreateStreamOnFileEx(filename.c_str(),
    STGM_READ, FILE_ATTRIBUTE_NORMAL, FALSE,
    nullptr,
    svg_stream.put()));
  
  winrt::check_hresult(d2d_dc->CreateSvgDocument(
    svg_stream.get(),
    D2D1::SizeF(1, 1),
    svg.put()));
  
  winrt::com_ptr<ID2D1SvgElement> root;
  svg->GetRoot(root.put());
  float tmp;
  winrt::check_hresult(root->GetAttributeValue(L"width", &tmp));
  svg_width = (int)tmp;
  winrt::check_hresult(root->GetAttributeValue(L"height", &tmp));
  svg_height = (int)tmp;
  return *this;
}

D2DSVG& D2DSVG::resize(int x, int y, int width, int height, float fill, float max_scale) {
  // Center 
  transform = D2D1::Matrix3x2F::Identity();
  transform = transform * D2D1::Matrix3x2F::Translation((width - svg_width) / 2.0f, (height - svg_height) / 2.0f);
  float h_scale = fill  * height / svg_height;
  float v_scale = fill * width / svg_width;
  used_scale = min(h_scale, v_scale);
  if (max_scale > 0) {
    used_scale = min(used_scale, max_scale);
  }
  transform = transform * D2D1::Matrix3x2F::Scale(used_scale, used_scale, D2D1::Point2F(width / 2.0f, height / 2.0f));
  transform = transform * D2D1::Matrix3x2F::Translation((float)x, (float)y);
  return *this;
}

D2DSVG& D2DSVG::render(ID2D1DeviceContext5* d2d_dc) {
  d2d_dc->SetTransform(transform);
  d2d_dc->DrawSvgDocument(svg.get());
  d2d_dc->SetTransform(D2D1::Matrix3x2F::Identity());
  return *this;
}

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

RECT D2DOverlaySVG::get_thumbnail_rect(int window_cx, int window_cy, float scale) {
  if (thumbnail_bottom_right.x == 0 && thumbnail_bottom_right.y == 0)
    return {};
  int thumbnail_scaled_rect_width = thumbnail_scaled_rect.right - thumbnail_scaled_rect.left;
  int thumbnail_scaled_rect_heigh = thumbnail_scaled_rect.bottom - thumbnail_scaled_rect.top;
  if (thumbnail_scaled_rect_heigh == 0 || thumbnail_scaled_rect_width == 0 ||
    window_cx == 0 || window_cy == 0) {
    return {};
  }
  float scale_h = scale * thumbnail_scaled_rect_width / window_cx;
  float scale_v = scale * thumbnail_scaled_rect_heigh / window_cy;
  float use_scale = min(scale_h, scale_v);
  RECT thumb_rect;
  thumb_rect.left = thumbnail_scaled_rect.left + (int)(thumbnail_scaled_rect_width - use_scale * window_cx) / 2;
  thumb_rect.right = thumbnail_scaled_rect.right - (int)(thumbnail_scaled_rect_width - use_scale * window_cx) / 2;
  thumb_rect.top = thumbnail_scaled_rect.top + (int)(thumbnail_scaled_rect_heigh - use_scale * window_cy) / 2;
  thumb_rect.bottom = thumbnail_scaled_rect.bottom - (int)(thumbnail_scaled_rect_heigh - use_scale * window_cy) / 2;
  return thumb_rect;
}

D2DOverlaySVG& D2DOverlaySVG::toggle_window_group(bool active) {
  if (window_group)
    window_group->SetAttributeValue(L"fill-opacity", active ? 1.0f : 0.3f);
  return *this;
}

void D2DWindow::init() {
  std::unique_lock<std::mutex> lock(mutex);
  // D2D1Factory is independent from the device, no need to recreate it if
  // we need to recreate the device.
  if (!d2d_factory) {
    D2D1_FACTORY_OPTIONS options = { D2D1_DEBUG_LEVEL_INFORMATION };
    winrt::check_hresult(D2D1CreateFactory(
      D2D1_FACTORY_TYPE_MULTI_THREADED,
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
  dxgi_device = nullptr;
  winrt::check_hresult(d3d_device->QueryInterface(
    __uuidof(dxgi_device),
    dxgi_device.put_void()));
  dxgi_factory = nullptr;
  winrt::check_hresult(CreateDXGIFactory2(
    0, // DXGI_CREATE_FACTORY_DEBUG for extra output, but might crash in releases
    __uuidof(dxgi_factory),
    dxgi_factory.put_void()));
  d2d_device = nullptr;
  winrt::check_hresult(d2d_factory->CreateDevice(dxgi_device.get(), d2d_device.put()));
  d2d_dc = nullptr;
  winrt::check_hresult(d2d_device->CreateDeviceContext(
    D2D1_DEVICE_CONTEXT_OPTIONS_NONE,
    d2d_dc.put()));
  landscape.load(L"svgs\\overlay.svg", d2d_dc.get())
           .find_thumbnail(L"path-1")
           .find_window_group(L"Group-1");
  portrait.load(L"svgs\\overlay_portrait.svg", d2d_dc.get())
           .find_thumbnail(L"path-1")
           .find_window_group(L"Group-1");
  no_active.load(L"svgs\\no_active_window.svg", d2d_dc.get());
  arrows.resize(10);
  for (unsigned i = 0; i < arrows.size(); ++i) {
    arrows[i].load(L"svgs\\" + std::to_wstring((i + 1) % 10) + L".svg", d2d_dc.get());
  }
}

void D2DWindow::resize() {
  std::unique_lock<std::mutex> lock(mutex);
  auto window_rect = *get_window_pos(hwnd);
  hwnd_rect.left = (float)window_rect.left;
  hwnd_rect.top = (float)window_rect.top;
  hwnd_rect.bottom = (float)window_rect.bottom;
  hwnd_rect.right = (float)window_rect.right;
  window_width = window_rect.right - window_rect.left;
  window_height = window_rect.bottom - window_rect.top;
  if (window_width == 0 || window_height == 0)
    return;
  DXGI_SWAP_CHAIN_DESC1 sc_description = {};
  sc_description.Format = DXGI_FORMAT_B8G8R8A8_UNORM;
  sc_description.BufferUsage = DXGI_USAGE_RENDER_TARGET_OUTPUT;
  sc_description.SwapEffect = DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL;
  sc_description.BufferCount = 2;
  sc_description.SampleDesc.Count = 1;
  sc_description.AlphaMode = DXGI_ALPHA_MODE_PREMULTIPLIED;
  sc_description.Width = window_width;
  sc_description.Height = window_height;
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

  float no_active_scale;
  if (window_width > window_height) {
    use_overlay = &landscape;
    no_active_scale = 0.3f;
  } else {
    use_overlay = &portrait;
    no_active_scale = 0.5f;
  }
  use_overlay->resize(0, 0, window_width, window_height, 0.95f);
  auto thumb_no_active_rect = use_overlay->get_thumbnail_rect(no_active.width(), no_active.height(), no_active_scale);
  no_active.resize(thumb_no_active_rect.left,
                   thumb_no_active_rect.top,
                   thumb_no_active_rect.right - thumb_no_active_rect.left,
                   thumb_no_active_rect.bottom - thumb_no_active_rect.top,
                   1.0f);
}
D2DSVG& D2DSVG::toggle_element(const wchar_t* id, bool visible) {
  winrt::com_ptr<ID2D1SvgElement> element;
  if (svg->FindElementById(id, element.put()) != S_OK)
    return *this;
  element->SetAttributeValue(L"display", visible ? D2D1_SVG_DISPLAY::D2D1_SVG_DISPLAY_INLINE : D2D1_SVG_DISPLAY::D2D1_SVG_DISPLAY_NONE);
  return *this;
}

void render_arrow(D2DSVG& arrow, TasklistButton& button, D2D1_RECT_F window, float max_scale, ID2D1DeviceContext5* d2d_dc) {
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
  } else {
    // same as above - make room for the hidden arrow
    auto render_arrow_height = button.height * 1.2f; 
    auto render_arrow_width = (int)(render_arrow_height / arrow_ratio);
    arrow.resize(dx == -1 ? button.x - render_arrow_width : button.x + button.width,
                 button.y + (button.height - render_arrow_height) / 2,
                 render_arrow_width, render_arrow_height, 0.95f, max_scale)
         .render(d2d_dc);
  }
}

bool D2DWindow::show_thumbnail() {
  if (!thumbnail)
    return false;
  SIZE thumb_size;
  if (DwmQueryThumbnailSourceSize(thumbnail, &thumb_size) != S_OK)
    return false;
  DWM_THUMBNAIL_PROPERTIES thumb_properties;
  thumb_properties.dwFlags = DWM_TNP_SOURCECLIENTAREAONLY | DWM_TNP_VISIBLE | DWM_TNP_RECTDESTINATION;
  thumb_properties.fSourceClientAreaOnly = FALSE;
  thumb_properties.fVisible = TRUE;
  thumb_properties.rcDestination = use_overlay->get_thumbnail_rect(thumb_size.cx, thumb_size.cy, 0.99f);
  if (thumb_properties.rcDestination.bottom == 0)
    return false;
  if (DwmUpdateThumbnailProperties(thumbnail, &thumb_properties) != S_OK)
    return false;
  return true;
}
void D2DWindow::render() {
  std::unique_lock<std::mutex> lock(mutex);
  if (!d2d_bitmap)
    return;
  d2d_dc->BeginDraw();
  d2d_dc->Clear();
  // Draw background
  winrt::com_ptr<ID2D1SolidColorBrush> brush;
  D2D1_COLOR_F const brushColor = D2D1::ColorF(1.0f, 1.0f, 1.0f, 0.8f);
  winrt::check_hresult(d2d_dc->CreateSolidColorBrush(brushColor, brush.put()));
  D2D1_RECT_F background_rect = {};
  background_rect.bottom = window_height;
  background_rect.right = window_width;
  d2d_dc->FillRectangle(background_rect, brush.get());
  // Draw SVG
  use_overlay->render(d2d_dc.get());
  if (show_thumbnail()) {
    use_overlay->toggle_window_group(true);
  } else {
    use_overlay->toggle_window_group(false);
    no_active.render(d2d_dc.get());
  }

  for (auto&& button : tasklist.get_buttons()) {
    if ((unsigned)button.keynum - 1 >= arrows.size())
      continue;
    render_arrow(arrows[button.keynum - 1], button, hwnd_rect, use_overlay->get_scale(), d2d_dc.get());
  }
  winrt::check_hresult(d2d_dc->EndDraw());
  winrt::check_hresult(dxgi_swap_chain->Present(1, 0));
  winrt::check_hresult(composition_device->Commit());
}

D2DWindow::~D2DWindow() {
  hide();
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
  case WM_MOVE:
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
