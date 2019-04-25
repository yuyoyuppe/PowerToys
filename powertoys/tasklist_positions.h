#pragma once
#include <vector>
#include <unordered_set>
#include <string>
#include <Windows.h>
#include <oleacc.h>

struct TasklistButton {
  std::wstring name;
  long x, y, width, height, keynum;
};

class Tasklist {
public:
  void update();
  std::vector<TasklistButton> get_buttons() const;
private:
  bool is_pinned(const std::wstring& name) const;
  std::vector<TasklistButton> assign_keynums(std::vector<TasklistButton> buttons) const;
  winrt::com_ptr<IAccessible> tasklist;
  bool labels_hidden;
  std::unordered_set<std::wstring> pinned_shortcuts;
  std::vector<char> pinned_registry;
};
