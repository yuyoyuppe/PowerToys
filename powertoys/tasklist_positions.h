#pragma once
#include <vector>
#include <Windows.h>

struct TasklistButton {
  long x, y, width, height, keynum;
  std::wstring name;
};
std::vector<TasklistButton> get_tasklist_buttons_positions();
