#include "pch.h"
#include "tasklist_positions.h"
#include <oleacc.h>
#pragma comment(lib, "oleacc.lib")

std::vector<TasklistButton> get_tasklist_buttons_positions() {
  std::vector<TasklistButton> rects;
  try {
    // It looks like this can ocasionally fail. In such cases just return empty list.
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
      if (left != 0 || top != 0 || width != 0 || height != 0) {
        apps_list = std::move(element);
        break;
      }
    }
    if (!apps_list)
      return rects;
    // Get positions of all buttons
    winrt::check_hresult(apps_list->get_accChildCount(&child_count));
    std::vector<TasklistButton> buttons;
    long min_width = -1, min_height = -1;
    for (int i = 0; i < child_count; ++i) {
      VARIANT cid, role;
      cid.vt = VT_I4;
      cid.lVal = i + 1;
      winrt::check_hresult(apps_list->get_accRole(cid, &role));
      if (role.vt != VT_I4 || role.lVal != ROLE_SYSTEM_PUSHBUTTON)
        continue;
      TasklistButton button;
      winrt::check_hresult(apps_list->accLocation(&button.x, &button.y, &button.width, &button.height, cid));
      buttons.push_back(button);
      if (button.width != 0 && (min_width == -1 || button.width < min_width))
        min_width = button.width;
      if (button.height != 0 && (min_height == -1 || button.height < min_height))
        min_height = button.height;
    }
    // Get names of pinned windows
    //auto pinned_list = FindFirstFile("%APPDATA%\\Microsoft\\Internet Explorer\\Quick Launch\\User Pinned\\TaskBar")
    int last_x = -1, last_y = -1;
    int keynum = 0;
    for (auto&& button : buttons) {
      if (button.width != 0 && button.height != 0) {
        if (last_x == -1) {
          last_x = button.x;
          last_y = button.y;
        }
        if (button.x < last_x || button.y < last_y) {
          // Ignore second row
          break;
        }
      }
      if (button.width == 0 || button.height == 0) {
        ++keynum;
        continue;
      }
      if (button.width == min_width && button.height == min_height) {
        ++keynum;
      }
      button.keynum = keynum;
      if (keynum == 11)
        break;
      if (rects.empty() || rects.back().keynum != keynum)
        rects.push_back(button);
    }
  } catch (...) {
    // Nothing to do, just return empty list
    rects.clear();
  }
  return rects;
}
