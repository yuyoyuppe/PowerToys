#include <Windows.h>
#include <stdexcept>

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
  HHOOK handle = NULL;
  try {
#ifdef _WIN64
    // Different named dlls are needed to hook processes from both architectures.
    HMODULE dll = LoadLibrary("powertoys_hook.dll");
#else
    HMODULE dll = LoadLibrary("powertoys_hook_32.dll");
#endif
    if (dll == NULL) {
      MessageBox(NULL, "Couldn't load powertoys_hook.dll", "Error", MB_OK | MB_ICONERROR);
      return -1;
    }
    HOOKPROC addr = (HOOKPROC)GetProcAddress(dll, "intercept_move_window_message");
    if (addr == NULL) {
      MessageBox(NULL, "Couldn't find leconnect functions in powertoys_hook.dll", "Error", MB_OK | MB_ICONERROR);
      return -1;
    }
    // Set the Windows hook to receive PostMessage from the powertoys executable.
    handle = SetWindowsHookEx(WH_GETMESSAGE, addr, dll, 0);
    if (handle == NULL) {
      MessageBox(NULL, "Couldn't hook powertoys_hook.dll", "Error", MB_OK | MB_ICONERROR);
      return -1;
    }
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0) > 0) {
      TranslateMessage(&msg);
      DispatchMessage(&msg);
    }
    int result = static_cast<int>(msg.wParam);
    UnhookWindowsHookEx(handle);
    return result;
  }
  catch (std::runtime_error err) {
    if (handle != NULL) {
      UnhookWindowsHookEx(handle);
    }
    MessageBox(NULL, err.what(), "Error", MB_OK | MB_ICONERROR);
    return -1;
  }
}
