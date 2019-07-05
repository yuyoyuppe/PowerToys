#pragma once

class Trace {
public:
  static void RegisterProvider();
  static void UnregisterProvider();
  static void EventShow();
  static void EventHide();
  static void ActionMaximize();
  static void ActionRestore();
};
