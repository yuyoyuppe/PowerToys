#include <windows.h>
#include <chrono>
using stdclock = std::chrono::system_clock;
#include <condition_variable>
#include <thread>
#include <ShellScalingApi.h>
#pragma comment(lib, "shcore.lib")

#include "popup_window.h"

PopupWindow* window;
HHOOK keyboard_hook_handle;
bool start_pressed = false;
stdclock::time_point start_press_timestamp;
std::mutex start_timer_mutex;
std::condition_variable start_timer_cv;
LRESULT CALLBACK keyboard_hook(int nCode, WPARAM wParam, LPARAM lParam) {
  auto kb_hook = reinterpret_cast<KBDLLHOOKSTRUCT*>(lParam);
  if (kb_hook->vkCode == VK_LWIN) {
    if (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN) {
      std::unique_lock<std::mutex> lock(start_timer_mutex);
      start_pressed = true;
      start_press_timestamp = stdclock::now();
      lock.unlock();
      start_timer_cv.notify_one();
    }
    else {
      start_pressed = false;
      window->hide();
    }
  }
  return CallNextHookEx(keyboard_hook_handle, nCode, wParam, lParam);
}

void start_timer_thread() {
  using namespace std::chrono_literals;
  while (true) {
    std::unique_lock<std::mutex> lock(start_timer_mutex);
    start_timer_cv.wait(lock, [] { return start_pressed; });
    lock.unlock();
    std::this_thread::sleep_for(300ms);
    if (start_pressed && stdclock::now() - start_press_timestamp > 300ms) {
      window->show(200, 200, 200, 30);
    }
  }
}



int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
  // We will handle scaling ourselfs
  SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE);
  window = new PopupWindow([](HWND hwnd) {
    PAINTSTRUCT ps;
    HDC hdc = BeginPaint(hwnd, &ps);
    TextOut(hdc, 10, 5, "content here", 15);
    EndPaint(hwnd, &ps);
    return 0;
  });
  
  std::thread start_thread_handle(start_timer_thread);
  start_thread_handle.detach();
  
  // Register hook for detecting Start menu press
  keyboard_hook_handle = SetWindowsHookEx(WH_KEYBOARD_LL,
                                          keyboard_hook,
                                          hInstance,
                                          NULL);
  if (keyboard_hook_handle == NULL) {
    MessageBox(NULL, "Cannot install hook", "ERROR!", MB_ICONERROR);
    return 0;
  }
  
  MSG msg;
  while (GetMessage(&msg, NULL, 0, 0)) {
    TranslateMessage(&msg);
    DispatchMessage(&msg);
  }
  return msg.wParam;
}