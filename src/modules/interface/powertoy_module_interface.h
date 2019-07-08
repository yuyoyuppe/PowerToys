#pragma once

/*
  DLL Interface for PowerToys. The powertoy_create() (see below) must return
  an object that implements this interface.

  See src/modules/example_powertoy for simple, noop, PowerToy implementation.

  The PowerToys runner will, for each PowerToy DLL:
    - load the DLL,
    - call powertoy_create() to create the PowerToy.

  On the received object, the runner will:
    - call get_name() to get the name of the PowerToy,
    - call get_config() to get the available configuration settings,
    - call get_events() to get the list of the events the PowerToy wants to subscribe to,
    - call enable() to initialize the PowerToy.

  While running, the runner might call disable()/enable(). It can also call
  set_config() to set various settings, anywhere between create_powertoy() and
  destroy().

  When terminating, the runner will:
    - call disable(),
    - call destroy() which should free all the memory and delete the PowerToy object,
    - unload the DLL.
 */

class PowertoyModuleIface {
public:
  /* Returns the name of the PowerToy, this will be cached by the runner. */
  virtual const wchar_t* get_name() = 0;
  /* Returns a null-terminated table of the names of the events the PowerToy wants to 
     subscribe to. Available events:
       * ll_keyboard

     A nullptr can be returned to signal that the PowerToy does not want to subscribe
     to any event.
  */
  virtual const wchar_t** get_events() = 0;
  /* Returns the available configuration settings. */
  virtual const wchar_t* get_config() = 0;
  /* Sets the configuration values. */
  virtual void set_config(const wchar_t* config) = 0;
  /* Enables the PowerToy. */
  virtual void enable() = 0;
  /* Disables the PowerToy, should free as much memory as possible. */
  virtual void disable() = 0;
  /* Handle event. Only the events the PowerToy subscribed to will be signaled.
     The data argument and return value meaning are event-specific:
       * ll_keyboard: see lowlevel_keyboard_evet_data.h.
  */
  virtual intptr_t signal_event(const wchar_t* name, intptr_t data) = 0;
  /* Destroy the PowerToy and free all memory. */
  virtual void destroy() = 0;
};

/*
  Typedef of the factory function that creates the PowerToy object.

  Must be exported by the DLL as powertoy_create(), e.g.:

  extern "C" __declspec(dllexport) PowertoyModuleIface* __cdecl powertoy_create()
  
  Called by the PowerToys runner to initialize each PowerToy.
  It will be called only once before a call to destroy() method is made.

  Returned PowerToy should be in disabled state. The runner will call
  the enable() method to start the PowerToy.

  In case of errors return nullptr.
*/
typedef PowertoyModuleIface* (__cdecl *powertoy_create_func)();
