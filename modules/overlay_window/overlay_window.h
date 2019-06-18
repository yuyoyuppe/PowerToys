#pragma once
#include "powertoys/powertoy_module_interface.h"
#include "powertoys/lowlevel_keyboard_event.h"
#include "d2d_overlay_window.h"

// We support only one instance of the overlay
extern class OverlayWindow* instance;

class TargetState;

class OverlayWindow : public PowertoyModuleIface {
public:
  OverlayWindow();
  virtual const wchar_t* get_name() override;
  virtual const wchar_t** get_events() override;
  virtual const wchar_t* get_config() override;
  virtual void set_config(const wchar_t* config) override;
  virtual void enable();
  virtual void disable();
  virtual intptr_t signal_event(const wchar_t* name, intptr_t data)  override;

  void on_held();
  void on_held_press(DWORD vkCode);
  void was_hidden();

  virtual void destroy() override;
private:
  TargetState* target_state;
  D2DOverlayWindow *winkey_popup;
  HWND desktop, shell;
  HWND active_window;
};
