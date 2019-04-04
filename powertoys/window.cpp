#include "window.h"
#include <exception>

std::recursive_mutex Window::static_mutex;
bool Window::window_class_initialized;
std::unordered_map<HWND, PaintProc> Window::paint_procedures;

/*
  Class for creating and displaying windows. Used params:
    WS_EX_TOOLWINDOW - window wont appear in Alt-Tab and on the taskbar
    WS_EX_TOPMOST - window should be on top
    WS_EX_LAYERED - can be made transparent by a call to SetLayeredWindowAttributes
    WS_EX_TRANSPARENT - window will be "clickthrough", all clicks will go to the
                        underlying window
    WS_POPUP - minimal window

  We keep hash map of WM_PAINT handlers for each window (paint_procedures). This
  way we can easily substitute paint handler to customize window content. We can
  even do that when resizing the window.
*/
Window::Window(PaintProc paint_proc) : fade(*this) {
  static const char* class_name = "PToyPopup";
  std::lock_guard<std::recursive_mutex> lock(static_mutex);
  if (!window_class_initialized) {
    WNDCLASSEX wnd_class;
    wnd_class.cbSize = sizeof(WNDCLASSEX);
    wnd_class.style = CS_SAVEBITS;
    wnd_class.lpfnWndProc = Window::window_proc;
    wnd_class.cbClsExtra = 0;
    wnd_class.cbWndExtra = 0;
    wnd_class.hInstance = GetModuleHandle(NULL);
    wnd_class.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    wnd_class.hCursor = LoadCursor(NULL, IDC_ARROW);
    wnd_class.hbrBackground = (HBRUSH)(COLOR_BACKGROUND + 1);
    wnd_class.lpszMenuName = NULL;
    wnd_class.lpszClassName = class_name;
    wnd_class.hIconSm = LoadIcon(NULL, IDI_APPLICATION);
    if (!RegisterClassEx(&wnd_class))
      throw std::runtime_error("Cannot register window class");
    window_class_initialized = true;
  }
  hwnd = CreateWindowEx(WS_EX_TOOLWINDOW | WS_EX_TOPMOST | WS_EX_LAYERED | WS_EX_TRANSPARENT,
                        class_name, class_name,
                        WS_POPUP,
                        CW_USEDEFAULT,
                        CW_USEDEFAULT, 240, 120,
                        NULL, NULL,
                        GetModuleHandle(NULL),
                        NULL);
  if (hwnd == NULL)
    throw std::runtime_error("Cannot create window");
  paint_procedures.emplace(hwnd, paint_proc);
}

void Window::set_transparency(double alpha) {
  SetLayeredWindowAttributes(hwnd, 0, (int)(255 * alpha), LWA_ALPHA);
}

void Window::show(int x_pos, int y_pos, int x_size, int y_size, PaintProc paint_proc) {
  paint_procedures[hwnd] = paint_proc;
  show(x_pos, y_pos, x_size, y_size);
}

void Window::show(int x_pos, int y_pos, int x_size, int y_size) {
  SetWindowPos(hwnd, HWND_TOPMOST, x_pos, y_pos, x_size, y_size, 0);
  HRGN elipse = CreateRoundRectRgn(0, 0, x_size, y_size, 15, 15);
  SetWindowRgn(hwnd, elipse, TRUE);
  show();
}

void Window::show() {
  ShowWindow(hwnd, SW_SHOWNA);
  UpdateWindow(hwnd);
}

void Window::hide() {
  ShowWindow(hwnd, SW_HIDE);
}

void Window::fade_in() {
  fade.fade_in();
}

void Window::fade_out() {
  fade.fade_out();
}

Window::~Window() {
  std::lock_guard<std::recursive_mutex> lock(static_mutex);
  paint_procedures.erase(hwnd);
}

LRESULT Window::window_proc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
  switch (msg) {
  case WM_PAINT: {
    std::lock_guard<std::recursive_mutex> lock(static_mutex);
    std::unordered_map<HWND, PaintProc>::const_iterator iter = paint_procedures.find(hwnd);
    if (iter == end(paint_procedures))
      return 0;
    auto& paint_proc = iter->second;
    return paint_proc(hwnd);
  }
  case WM_CLOSE:
    DestroyWindow(hwnd);
    break;
  case WM_DESTROY:
    PostQuitMessage(0);
    break;
  default:
    return DefWindowProc(hwnd, msg, wParam, lParam);
  }
  return 0;
}

FadeWindow::FadeWindow(Window& window) : running(false), exit(false), window_ptr(&window) {
  thread = std::thread(&FadeWindow::thread_proc, this);
}

void FadeWindow::fade_in() {
  std::unique_lock<std::mutex> lock(mutex);
  if (!running) {
    next = 0;
  }
  delta = 0.05;
  running = true;
  lock.unlock();
  cv.notify_one();
}

void FadeWindow::fade_out() {
  std::unique_lock<std::mutex> lock(mutex);
  if (!running) {
    next = 0.7;
  }
  delta = -0.05;
  running = true;
  lock.unlock();
  cv.notify_one();
}

FadeWindow::~FadeWindow() {
  std::unique_lock<std::mutex> lock(mutex);
  exit = true;
  lock.unlock();
  cv.notify_one();
  thread.join();
}

void FadeWindow::thread_proc() {
  std::unique_lock<std::mutex> lock(mutex);
  while (!exit) {
    cv.wait(lock, [&] { return running || exit; });
    if (exit)
      return;
    while (!exit) {
      if (next < 0) {
        window_ptr->hide();
        break;
      }
      if (next > 0.7) {
        window_ptr->set_transparency(0.7);
        break;
      }
      window_ptr->set_transparency(next);
      next += delta;
      std::this_thread::sleep_for(std::chrono::milliseconds(16));
    }
    running = false;
  }
}
