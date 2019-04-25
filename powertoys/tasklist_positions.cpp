#include "pch.h"
#include "tasklist_positions.h"
#pragma comment(lib, "oleacc.lib")

void Tasklist::update() {
  // Get HWND of the tasklist
  auto tasklist_hwnd = FindWindow("Shell_TrayWnd", nullptr);
  if (!tasklist_hwnd) return;
  tasklist_hwnd = FindWindowEx(tasklist_hwnd, 0, "ReBarWindow32", nullptr);
  if (!tasklist_hwnd) return;
  tasklist_hwnd = FindWindowEx(tasklist_hwnd, 0, "MSTaskSwWClass", nullptr);
  if (!tasklist_hwnd) return;
  tasklist_hwnd = FindWindowEx(tasklist_hwnd, 0, "MSTaskListWClass", nullptr);
  if (!tasklist_hwnd) return;
  tasklist = nullptr;
  if (AccessibleObjectFromWindow(tasklist_hwnd, OBJID_WINDOW, IID_IAccessible, tasklist.put_void()) < 0)
    return;
  // Get localized names of pinned shortcuts
  pinned_shortcuts.clear();
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
      pinned_shortcuts.insert(info.szDisplayName);
    } while (FindNextFileW(pinned_list, &file_data) != 0);
  }
  // Get registry value of Favorites - it contains pinned store apps
  pinned_registry.clear();
  DWORD stored_type, data_size = 0;
  if (RegGetValue(HKEY_CURRENT_USER,
                  R"(Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband)",
                  "FavoritesResolve",
                  RRF_RT_REG_BINARY,
                  &stored_type,
                  NULL,
                  &data_size) == ERROR_SUCCESS) {
    pinned_registry.resize(data_size);
    RegGetValue(HKEY_CURRENT_USER,
                R"(Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband)",
                "FavoritesResolve",
                RRF_RT_REG_BINARY,
                &stored_type,
                pinned_registry.data(),
                &data_size);
  }
  // Check if labels are always hidden
  labels_hidden = false;
  DWORD reg_value;
  data_size = sizeof(DWORD);
  auto res = RegGetValue(HKEY_CURRENT_USER,
                         R"(Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced)",
                         "TaskbarGloming",
                         RRF_RT_REG_DWORD,
                         &stored_type,
                         &reg_value,
                         &data_size);
  if (res == ERROR_SUCCESS && reg_value == 0) // combinig icons is disabled
    return;
  res = RegGetValue(HKEY_CURRENT_USER,
                    R"(Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced)",
                    "TaskbarGlomLevel",
                    RRF_RT_REG_DWORD,
                    &stored_type,
                    &reg_value,
                    &data_size);
  if (res != ERROR_SUCCESS || reg_value != 0)
    return;
  labels_hidden = true;
}

bool Tasklist::is_pinned(const std::wstring& name) const {
  if (pinned_shortcuts.find(name) != pinned_shortcuts.end())
    return true;
  std::string_view registry(pinned_registry.data(), pinned_registry.size());
  std::string_view item_name((const char*)name.data(), name.size() * 2);
  auto pos = registry.find(item_name);
  return pos != -1;
  if (pos == -1)
    return false;
}

std::vector<TasklistButton> Tasklist::assign_keynums(std::vector<TasklistButton> buttons) const {
  std::vector<TasklistButton> rects;
  int keynum = 0;
  long last_x = -1, last_y = -1;
  if (labels_hidden) {
    // simplest case - all icons are combined
    for (auto&& button : buttons) {
      if (button.width == 0 || button.height == 0)
        continue;
      if (rects.empty() || last_x != button.x || last_y != button.y) {
        button.keynum = ++keynum;
        rects.push_back(button);
        last_x = button.x;
        last_y = button.y;
        if (keynum == 10)
          break;
      }
    }
  } else {
    // labels are visible
    //  try some heurystics to get correct keynums
    bool next_is_next_keynum = false;
    bool next_keynum = false;
    for (auto&& button : buttons) {
      if (button.width == 0 || button.height == 0) {
        next_keynum = true;
        next_is_next_keynum = false;
        continue;
      }
      if (is_pinned(button.name)) {
        next_is_next_keynum = false;
        next_keynum = true;
      }
      if (button.x == last_x && button.y == last_y) {
        next_is_next_keynum = true;
        next_keynum = true;
        continue;
      } else if (next_is_next_keynum) {
        next_is_next_keynum = false;
        next_keynum = true;
      }
      last_x = button.x;
      last_y = button.y;

      if (next_keynum) {
        button.keynum = ++keynum;
        rects.push_back(button);
        if (keynum == 10)
          break;
      }  
    }
  }
  return rects;
}

std::vector<TasklistButton> Tasklist::get_buttons() const {
  // It looks like this can ocasionally fail. In such cases just return empty list.
  if (!tasklist)
    return {};
  VARIANT children[256];
  long child_count;
  if (AccessibleChildren(tasklist.get(), 0, 256, children, &child_count) < 0)
    return {};
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
    return {};
  // Get positions of all buttons
  if (apps_list->get_accChildCount(&child_count) < 0)
    return {};
  std::vector<TasklistButton> buttons;
  int last_x = -1, last_y = -1;
  for (int i = 0; i < child_count; ++i) {
    VARIANT cid, role;
    cid.vt = VT_I4;
    cid.lVal = i + 1;
    if (apps_list->get_accRole(cid, &role) < 0 || role.vt != VT_I4 || (role.lVal != ROLE_SYSTEM_PUSHBUTTON && role.lVal != ROLE_SYSTEM_BUTTONMENU))
      continue;
    TasklistButton button;
    if (apps_list->accLocation(&button.x, &button.y, &button.width, &button.height, cid) < 0)
      continue;
    if (button.width != 0 && button.height != 0) {
      if (last_x != -1 && (button.x < last_x || button.y < last_y)) {
        // Ignore second row
        break;
      }
      last_x = button.x;
      last_y = button.y;
    }
    BSTR name;
    if (apps_list->get_accName(cid, &name) >= 0) {
      button.name = name;
      SysFreeString(name);
    }
    buttons.push_back(button);
  }
  return assign_keynums(std::move(buttons));
}
