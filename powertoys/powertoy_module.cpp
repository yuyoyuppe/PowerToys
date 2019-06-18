#include "pch.h"
#include "powertoy_module.h"
#include <algorithm>

PowertoysEvents& powertoys_events() {
  static PowertoysEvents powertoys_events;
  return powertoys_events;
}

void PowertoysEvents::register_receiver(const std::wstring & event, PowertoyModuleIface * module) {
  std::unique_lock<std::recursive_mutex> lock(mutex);
  receivers[event].push_back(module);
}

void PowertoysEvents::unregister_receiver(PowertoyModuleIface* module) {
  std::unique_lock<std::recursive_mutex> lock(mutex);
  for (auto& kv : receivers) {
    auto& vec = kv.second;
    vec.erase(remove(begin(vec), end(vec), module), end(vec));
  }
}

intptr_t PowertoysEvents::signal_event(const std::wstring & event, intptr_t data) {
  intptr_t rvalue = 0;
  std::unique_lock<std::recursive_mutex> lock(mutex);
  if (auto it = receivers.find(event); it != end(receivers)) {
    for (auto& module : it->second) {
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
  return PowertoyModule(create(), handle);
}
