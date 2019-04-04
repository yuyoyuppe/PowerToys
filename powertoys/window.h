#pragma once
#include <functional>
#include <unordered_map>
#include <mutex>
#include <thread>
#include <condition_variable>
#include <windows.h>

typedef std::function<LRESULT(HWND hwnd)> PaintProc;

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
  Window(PaintProc paint_proc = [](HWND) { return 0; });
  void set_transparency(double alpha);
  void show(int x_pos, int y_pos, int x_size, int y_size, PaintProc paint_proc);
  void show(int x_pos, int y_pos, int x_size, int y_size);
  void show();
  void hide();
  void fade_in();
  void fade_out();
  ~Window();
  HWND hwnd;
private:
  FadeWindow fade;
  static LRESULT CALLBACK window_proc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
  static std::recursive_mutex static_mutex;
  static bool window_class_initialized;
  static std::unordered_map<HWND, PaintProc> paint_procedures;
};
