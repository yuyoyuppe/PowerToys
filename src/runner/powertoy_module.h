#pragma once
#include <interface/powertoy_module_interface.h>
#include <string>
#include <memory>
#include <mutex>
#include <vector>
#include <functional>

class PowertoyModule;

class PowertoysEvents {
public:
  void register_receiver(const std::wstring& event, PowertoyModuleIface* module);
  void unregister_receiver(PowertoyModuleIface* module);
  intptr_t signal_event(const std::wstring& event, intptr_t data);
private:
  std::recursive_mutex mutex;
  std::unordered_map<std::wstring, std::vector<PowertoyModuleIface*>> receivers;
};

PowertoysEvents& powertoys_events();

struct PowertoyModuleDeleter {
  void operator()(PowertoyModuleIface* module) {
    if (module) {
      powertoys_events().unregister_receiver(module);
      module->disable();
      module->destroy();
    }
  }
};

struct PowertoyModuleDLLDeleter {
  using pointer = HMODULE;
  void operator()(HMODULE  handle) {
    FreeLibrary(handle);
  }
};


class PowertoyModule {
public:
  PowertoyModule(PowertoyModuleIface* module, HMODULE  handle) : handle(handle), module(module) {
    name = module->get_name();
    config = module->get_config();
    auto want_signals = module->get_events();
    if (want_signals) {
      for (; *want_signals; ++want_signals) {
        powertoys_events().register_receiver(*want_signals, module);
      }
    }
  }
  const std::wstring& get_name() const {
    return name;
  }
  const std::wstring& get_confing() const {
    return config;
  }
  void set_config(const std::wstring& config) {
    module->set_config(config.c_str());
  }
  intptr_t signal_event(const std::wstring& signal_event, intptr_t data) {
    return module->signal_event(signal_event.c_str(), data);
  }
  void enable() {
    module->enable();
  }
  void disable() {
    module->disable();
  }
private:
  std::unique_ptr<HMODULE, PowertoyModuleDLLDeleter> handle;
  std::unique_ptr<PowertoyModuleIface, PowertoyModuleDeleter> module;
  std::wstring name;
  std::wstring config;
};


PowertoyModule load_powertoy(const std::wstring& filename);
std::unordered_map<std::wstring, PowertoyModule>& modules();