#pragma once
#include <functional>
#include <unordered_map>
#include <mutex>
#include <windows.h>

typedef std::function<LRESULT(HWND hwnd)> PaintProc;

class PopupWindow {
public:
  PopupWindow(PaintProc paint_proc = [](HWND) { return 0; });
  void set_transparency(double alpha);
  void show(int x_pos, int y_pos, int x_size, int y_size, PaintProc paint_proc);
  void show(int x_pos, int y_pos, int x_size, int y_size);
  void show();
  void hide();
  ~PopupWindow();
  HWND hwnd;
private:
  
  static LRESULT CALLBACK window_proc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
  static std::recursive_mutex static_mutex;
  static bool window_class_initialized;
  static std::unordered_map<HWND, PaintProc> paint_procedures;
};
