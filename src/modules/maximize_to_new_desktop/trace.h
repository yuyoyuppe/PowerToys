#pragma once

class Trace {
public:
  static void RegisterProvider();
  static void UnregisterProvider();
  static void EventShow();
  static void EventHide(const __int64 duration_ms);
  static void EventDesktopClosed();
  static void ActionMaximize();
  static void ActionRestore();
};
