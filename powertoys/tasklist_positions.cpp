#include "pch.h"
#include "tasklist_positions.h"
#include <oleacc.h>
#pragma comment(lib, "oleacc.lib")

std::vector<RECT> get_tasklist_buttons_positions() {
  std::vector<RECT> rects;
  auto tasklist_hwnd = FindWindow("Shell_TrayWnd", nullptr);
  if (!tasklist_hwnd)
    return rects;
  tasklist_hwnd = FindWindowEx(tasklist_hwnd, 0, "ReBarWindow32", nullptr);
  if (!tasklist_hwnd)
    return rects;
  tasklist_hwnd = FindWindowEx(tasklist_hwnd, 0, "MSTaskSwWClass", nullptr);
  if (!tasklist_hwnd)
    return rects;
  tasklist_hwnd = FindWindowEx(tasklist_hwnd, 0, "MSTaskListWClass", nullptr);
  if (!tasklist_hwnd)
    return rects;
  winrt::com_ptr<IAccessible> accessible;
  winrt::check_hresult(AccessibleObjectFromWindow(tasklist_hwnd, OBJID_WINDOW, IID_IAccessible, accessible.put_void()));
  VARIANT children[256];
  long child_count;
  winrt::check_hresult(AccessibleChildren(accessible.get(), 0, 256, children, &child_count));
  winrt::com_ptr<IAccessible> apps_list;
  // look for child with x/y set
  long left, top, width, height;
  for (int i = 0; i < child_count; ++i) {
    if (children[i].vt != VT_DISPATCH)
      continue;
    winrt::com_ptr<IDispatch> dispatch;
    dispatch.attach(children[i].pdispVal);
    winrt::com_ptr<IAccessible> element;
    winrt::check_hresult(dispatch->QueryInterface(IID_IAccessible, element.put_void()));
    VARIANT cid;
    cid.vt = VT_I4;
    cid.lVal = CHILDID_SELF;
    element->accLocation(&left, &top, &width, &height, cid);
    if (left != 0 && top != 0 && width != 0 && height != 0) {
      apps_list = std::move(element);
      break;
    }
  }
  if (!apps_list)
    return rects;
  // get positions
  winrt::check_hresult(AccessibleChildren(apps_list.get(), 0, 256, children, &child_count));
  for (int i = 0; i < child_count; ++i) {
    VARIANT cid;
    cid.vt = VT_I4;
    cid.lVal = i + 1;
    winrt::check_hresult(apps_list->accLocation(&left, &top, &width, &height, cid));
    if (left != 0 && top != 0 && width != 0 && height != 0) {
      RECT rect;
      rect.left = left;
      rect.top = top;
      rect.right = left + width;
      rect.bottom = top + height;
      rects.push_back(rect);
      if (rects.size() == 9)
        break;
    }
  }
  return rects;
}