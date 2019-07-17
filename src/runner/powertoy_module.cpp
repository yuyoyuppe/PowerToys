#include "pch.h"
#include "powertoy_module.h"
#include "lowlevel_keyboard_event.h"
#include <algorithm>

std::unordered_map<std::wstring, PowertoyModule>& modules() {
  // make sure events are destroyed after modules
  powertoys_events();
  static std::unordered_map<std::wstring, PowertoyModule> modules;
  return modules;
}

PowertoyModule load_powertoy(const std::wstring& filename) {
  auto handle = LoadLibraryW(filename.c_str());
  if (!handle) {
    throw std::runtime_error("Cannot load " + std::string(begin(filename), end(filename)));
  }
  auto create = reinterpret_cast<powertoy_create_func>(GetProcAddress(handle, "powertoy_create"));
  if (!create) {
    FreeLibrary(handle);
    throw std::runtime_error("Cannot load factory function from " + std::string(begin(filename), end(filename)));
  }
  auto module = create();
  if (!module) {
    throw std::runtime_error("Cannot create module " + std::string(begin(filename), end(filename)));
  }
  return PowertoyModule(module, handle);
}
