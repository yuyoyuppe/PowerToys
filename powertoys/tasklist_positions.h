#pragma once
#include <vector>
#include <Windows.h>

struct TasklistButton {
  std::wstring name;
  long x, y, width, height, keynum, role_id;
};
std::vector<TasklistButton> get_tasklist_buttons_positions();
