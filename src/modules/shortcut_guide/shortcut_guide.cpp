#include "pch.h"
#include "shortcut_guide.h"
#include "target_state.h"
#include "trace.h"

OverlayWindow* instance = nullptr;

OverlayWindow::OverlayWindow() : winkey_popup(new D2DOverlayWindow()), target_state(new TargetState(300)) {
  winkey_popup->initialize();
  desktop = GetDesktopWindow();
  shell = GetShellWindow();
}

const wchar_t * OverlayWindow::get_name() {
  return L"Shortcut Guide";
}

const wchar_t ** OverlayWindow::get_events() {
  static const wchar_t* events[2] = { L"ll_keyboard", 0 };
  return events;
}

const wchar_t * OverlayWindow::get_config() {
  return L"";
}

void OverlayWindow::set_config(const wchar_t * config) { }

void OverlayWindow::enable() { }

void OverlayWindow::disable() { }

intptr_t OverlayWindow::signal_event(const wchar_t * name, intptr_t data) {
  if (wcscmp(name, L"ll_keyboard") == 0) {
    auto& event = *(reinterpret_cast<LowlevelKeyboardEvent*>(data));
    if (event.wParam == WM_KEYDOWN ||
        event.wParam == WM_SYSKEYDOWN ||
        event.wParam == WM_KEYUP ||
        event.wParam == WM_SYSKEYUP) {
      bool supress = target_state->signal_event(event.lParam->vkCode, 
                                                event.wParam == WM_KEYDOWN || event.wParam == WM_SYSKEYDOWN);
      return supress ? 1 : 0;
    }
  }
  return 0;
}

void OverlayWindow::on_held() {
  auto active_window = GetForegroundWindow();
  active_window = GetAncestor(active_window, GA_ROOT);
  if (active_window == desktop || active_window == shell) {
    active_window = nullptr;
  }
  auto window_styles = active_window ? GetWindowLong(active_window, GWL_STYLE) : 0;
  if ((window_styles & WS_CHILD) || (window_styles & WS_DISABLED)) {
    active_window = nullptr;
  }
  char class_name[256] = "";
  GetClassNameA(active_window, class_name, 256);
  if (strcmp(class_name, "SysListView32") == 0 ||
    strcmp(class_name, "WorkerW") == 0 ||
    strcmp(class_name, "Shell_TrayWnd") == 0 ||
    strcmp(class_name, "Shell_SecondaryTrayWnd") == 0) {
    active_window = nullptr;
  }
  winkey_popup->show(active_window);
}

void OverlayWindow::on_held_press(DWORD vkCode) {
  winkey_popup->animate(vkCode);
}

void OverlayWindow::was_hidden() {
  target_state->was_hiden();
}

void OverlayWindow::destroy() {
  winkey_popup->hide();
  target_state->exit();
  delete target_state;
  delete winkey_popup;
  delete this;
  instance = nullptr;
}
