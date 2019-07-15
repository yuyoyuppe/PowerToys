#include "pch.h"
#include "powertoy_module.h"
#include <algorithm>

std::unordered_map<std::wstring, PowertoyModule>& modules() {
  // make sure events are destroyed after modules
  powertoys_events();
  static std::unordered_map<std::wstring, PowertoyModule> modules;
  return modules;
}

PowertoysEvents& powertoys_events() {
  static PowertoysEvents powertoys_events;
  return powertoys_events;
}

void PowertoysEvents::register_receiver(const std::wstring & event, PowertoyModuleIface * module) {
  std::unique_lock lock(mutex);
  receivers[event].push_back(module);
}

void PowertoysEvents::unregister_receiver(PowertoyModuleIface* module) {
  std::unique_lock lock(mutex);
  for (auto& [key, value] : receivers) {
    value.erase(remove(begin(value), end(value), module), end(value));
  }
}

intptr_t PowertoysEvents::signal_event(const std::wstring & event, intptr_t data) {
  intptr_t rvalue = 0;
  std::unique_lock lock(mutex);
  if (auto it = receivers.find(event); it != end(receivers)) {
    for (auto& module : it->second) {
      if (module)
        rvalue |= module->signal_event(event.c_str(), data);
    }
  }
  return rvalue;
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
