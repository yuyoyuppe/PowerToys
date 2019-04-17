#pragma once
#include <vector>
#include <Windows.h>

struct TasklistButton {
  long x, y, width, height, keynum;
};
std::vector<TasklistButton> get_tasklist_buttons_positions();
