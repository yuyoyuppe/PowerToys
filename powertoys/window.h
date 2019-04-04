#pragma once
#include <functional>
#include <unordered_map>
#include <mutex>
#include <thread>
#include <condition_variable>
#include <windows.h>

typedef std::function<LRESULT(HWND hwnd, WPARAM wParam, LPARAM)> MsgProc;

class Window;

class FadeWindow {
public:
  FadeWindow(Window& window);
  void fade_in();
  void fade_out();
  ~FadeWindow();
private:
  Window* window_ptr;
  double next, delta;
  bool running, exit;
  std::thread thread;
  std::mutex mutex;
  std::condition_variable cv;
  void thread_proc();
};

class Window {
public:
  Window();
  Window& set_transparency(double alpha);
  Window& round_corners(int radius);
  Window& add_handler(UINT msg, MsgProc msg_proc);
  Window& show(int x_pos, int y_pos, int x_size, int y_size, MsgProc wmpaint_proc);
  Window& show(int x_pos, int y_pos, int x_size, int y_size);
  Window& show(RECT rect);
  Window& show(RECT, MsgProc wmpaint_proc);
  Window& show();
  Window& hide();
  Window& fade_in();
  Window& fade_out();
  ~Window();
private:
  int corner_radius;
  HWND hwnd;
  FadeWindow fade;
  static LRESULT CALLBACK window_proc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
  // Recursive mutex allows us to create/modify other Window instances from the message callbacks
  static std::recursive_mutex static_mutex;
  static bool window_class_initialized;
  static std::unordered_map<HWND, std::unordered_map<UINT, MsgProc>> msg_procedures;
};
