#pragma once
#include <interface/powertoy_module_interface.h>
#include <interface/lowlevel_keyboard_event_data.h>
#include "overlay_window.h"

// We support only one instance of the overlay
extern class OverlayWindow* instance;

class TargetState;

class OverlayWindow : public PowertoyModuleIface {
public:
  OverlayWindow();
  virtual const wchar_t* get_name() override;
  virtual const wchar_t** get_events() override;
  virtual bool get_config(const wchar_t** config) override;
  virtual void free_get_config(const wchar_t* config) override;

  virtual void set_config(const wchar_t* config) override;
  virtual void enable() override;
  virtual void disable() override;
  virtual bool is_enabled() override;
  virtual intptr_t signal_event(const wchar_t* name, intptr_t data)  override;

  int get_current_delay_setting();
  void set_current_delay_setting(int new_delay);

  void on_held();
  void on_held_press(DWORD vkCode);
  void was_hidden();

  virtual void destroy() override;
private:
  TargetState* target_state;
  D2DOverlayWindow *winkey_popup;
  HWND desktop, shell;
  HWND active_window;
  bool _enabled = false;

  int current_delay_setting = 900;
  void init_settings();
  void save_settings();
};
