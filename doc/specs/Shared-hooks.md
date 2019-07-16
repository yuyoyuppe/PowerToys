# Shared hooks

To minimize the performance impact on the machine only `runner` installs global hooks, passing the events to registered callbacks in each PowerToy module.

When a PowerToy module is loaded, the `runner` calls the [`get_events()`](/src/modules/interface/powertoy_module_interface.h#L40) method to get a NULL-terminated array of NULL-terminated strings with the names of the events that the PowerToy wants to subcribe to.

Events are signalled by the `runner` calling the [`signal_event(name, data)`](/src/modules/interface/powertoy_module_interface.h#L53) method of the PowerToy module. The `name` parameter contains the NULL-terminated name of the event. The `data` parameter and the method return value are specific for each event.

Currently supported hooks:
 * `"ll_keyboard"` - [Low Level Keyboard Hook](#low-level-keyboard-hook)

## Low Level Keyboard Hook

This event is signaled whenever user presses or releases a key on the keyboard. To subscribe to this event, add `"ll_keyboard"` to the table returned by the `get_events()` methods.

The PowerToys runner installs low-level keyboard hook using `SetWindowsHookEx(WH_KEYBOARD_LL, ...)`. See [this MSDN page](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms644985(v%3Dvs.85)) for details.

When a keyboard event is signaled and `ncCode` equals `HC_ACTION`, the `wParam` and `lParam` event parameters are passed to all subscribed clients in the [`LowlevelKeyboardEvent`](/src/modules/interface/lowlevel_keyboard_event_data.h#L38-L41) struct.

The `intptr_t data` event argument is a pointer to the `LowlevelKeyboardEvent` struct.

A non-zero return value from any of the subscribed PowerToys will cause the runner hook proc to return 1, thus swallowing the keyboard event.

Example usage, that makes Windows ignore the L key:

```c++
virtual const wchar_t** get_events() override {
  static const wchar_t* events[2] = { L"ll_keyboard",
                                      nullptr };
  return events;
}

virtual intptr_t signal_event(const wchar_t* name, intptr_t data) override {
  if (wcscmp(name, L"ll_keyboard") == 0) {
    auto& event = *(reinterpret_cast<LowlevelKeyboardEvent*>(data));
    // The L key has vkCode of 0x4C, see:
    // https://docs.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes
    if (event.wParam == WM_KEYDOWN && event.lParam->vkCode == 0x4C) {
      return 1;
    } else {
      return 0;
    }
  } else {
    return 0;
  }
}
```
