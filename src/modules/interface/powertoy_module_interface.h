#pragma once

class PowertoyModuleIface {
public:
  virtual const wchar_t* get_name() = 0;
  virtual const wchar_t** get_events() = 0;
  virtual const wchar_t* get_config() = 0;
  virtual void set_config(const wchar_t* config) = 0;
  virtual void enable() = 0;
  virtual void disable() = 0;
  virtual intptr_t signal_event(const wchar_t* name, intptr_t data) = 0;
  virtual void destroy() = 0;
};


typedef PowertoyModuleIface* (__cdecl *powertoy_create_func)();