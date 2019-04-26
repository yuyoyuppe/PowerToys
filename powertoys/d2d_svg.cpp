#include "pch.h"
#include "d2d_svg.h"

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
  float h_scale = fill * height / svg_height;
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

D2DSVG& D2DSVG::toggle_element(const wchar_t* id, bool visible) {
  winrt::com_ptr<ID2D1SvgElement> element;
  if (svg->FindElementById(id, element.put()) != S_OK)
    return *this;
  element->SetAttributeValue(L"display", visible ? D2D1_SVG_DISPLAY::D2D1_SVG_DISPLAY_INLINE : D2D1_SVG_DISPLAY::D2D1_SVG_DISPLAY_NONE);
  return *this;
}
