#include "pch.h"
#include "tasklist_positions.h"
#include <oleacc.h>
#pragma comment(lib, "oleacc.lib")

std::unordered_set<std::wstring> get_pinned_items() {
  std::unordered_set<std::wstring> rval;
  std::wstring links_folder;
  links_folder.resize(GetEnvironmentVariableW(L"APPDATA", nullptr, 0));
  GetEnvironmentVariableW(L"APPDATA", links_folder.data(), links_folder.length());
  links_folder.pop_back();
  links_folder.append(L"\\Microsoft\\Internet Explorer\\Quick Launch\\User Pinned\\TaskBar\\");
  // Get names of pinned windows
  WIN32_FIND_DATAW file_data;
  auto pinned_list = FindFirstFileW((links_folder + L"*.lnk").c_str(), &file_data);
  if (pinned_list != INVALID_HANDLE_VALUE) {
    do {
      // Get localized file name
      SHFILEINFOW info;
      SHGetFileInfoW((links_folder + file_data.cFileName).c_str(), FILE_ATTRIBUTE_NORMAL, &info, sizeof(info), SHGFI_DISPLAYNAME);
      rval.insert(info.szDisplayName);
    } while (FindNextFileW(pinned_list, &file_data) != 0);
  }
  return rval;
}

std::vector<TasklistButton> get_tasklist_buttons_positions() {
  std::vector<TasklistButton> rects;
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
  if (AccessibleObjectFromWindow(tasklist_hwnd, OBJID_WINDOW, IID_IAccessible, accessible.put_void()) < 0)
    return rects;
  VARIANT children[256];
  long child_count;
  if (AccessibleChildren(accessible.get(), 0, 256, children, &child_count) < 0)
    return rects;
  winrt::com_ptr<IAccessible> apps_list;
  // look for child with x/y set
  long left, top, width, height;
  for (int i = 0; i < child_count; ++i) {
    if (children[i].vt != VT_DISPATCH)
      continue;
    winrt::com_ptr<IDispatch> dispatch;
    dispatch.attach(children[i].pdispVal);
    winrt::com_ptr<IAccessible> element;
    if (dispatch->QueryInterface(IID_IAccessible, element.put_void()) < 0)
      continue;
    VARIANT cid;
    cid.vt = VT_I4;
    cid.lVal = CHILDID_SELF;
    if (element->accLocation(&left, &top, &width, &height, cid) < 0)
      continue;
    if (left != 0 || top != 0 || width != 0 || height != 0) {
      apps_list = std::move(element);
      break;
    }
  }
  if (!apps_list)
    return rects;
  // Get positions of all buttons
  if (apps_list->get_accChildCount(&child_count) < 0)
    return rects;
  std::vector<TasklistButton> buttons;
  for (int i = 0; i < child_count; ++i) {
    VARIANT cid, role;
    cid.vt = VT_I4;
    cid.lVal = i + 1;
    if (apps_list->get_accRole(cid, &role) < 0 || role.vt != VT_I4 || role.lVal != ROLE_SYSTEM_PUSHBUTTON)
      continue;
    TasklistButton button; 
    if (apps_list->accLocation(&button.x, &button.y, &button.width, &button.height, cid) < 0)
      continue;
    BSTR name;
    if (apps_list->get_accName(cid, &name) < 0)
      continue;
    button.name = name;
    SysFreeString(name);
    buttons.push_back(button);
  }
  auto pinned_items = get_pinned_items();
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
    if (pinned_items.find(button.name) != pinned_items.end()) {
      ++keynum;
    }
    button.keynum = keynum;
    if (keynum == 11)
      break;
    if (rects.empty() || rects.back().keynum != keynum)
    rects.push_back(button);
  }
  return rects;
}
