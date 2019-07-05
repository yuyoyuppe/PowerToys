#include "pch.h"
#include "tray_icon.h"
#include <ShellScalingApi.h>
#include "powertoy_module.h"
#include "lowlevel_keyboard_event.h"
#include <filesystem>
#include "trace.h"

#if _DEBUG && _WIN64
#include "unhandled_exception_handler.h"
#endif

void chdir_current_executable() {
  // Change current directory to the path of the executable.
  TCHAR executable_path[MAX_PATH];
  GetModuleFileName(NULL, executable_path, MAX_PATH);
  PathRemoveFileSpec(executable_path);
  if(!SetCurrentDirectory(executable_path)) {
    show_last_error_message(TEXT("Change Directory to Executable Path"), GetLastError());
  }
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
  #if _DEBUG && _WIN64
  //Global error handlers to diagnose errors.
  //We prefer this not not show any longer until there's a bug to diagnose.
  //init_global_error_handlers();
  #endif
  Trace::RegisterProvider();
  winrt::init_apartment();
  start_tray_icon();
  winrt::check_hresult(SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE));
  int result;
  try {
    chdir_current_executable();
    // Load Powertyos DLLS
    // For now only load know DLLs
    std::unordered_set<std::wstring> know_dlls = {
      L"example_powertoy.dll",
      L"shortcut_guide.dll",
      L"maximize_to_new_desktop.dll" 
    };
    for (auto& file : std::filesystem::directory_iterator(TEXT("modules/"))) {
      if (file.path().extension() != L".dll")
        continue;
      if (know_dlls.find(file.path().filename()) == know_dlls.end())
        continue;
      try {
        auto module = load_powertoy(file.path().wstring());
        modules().emplace(module.get_name(), std::move(module));
      } catch (...) { }
    } 
    // Start our events providers
    start_lowlevel_keyboard_hook();
    // Start all the powertoys
    for (auto& [name, powertoy] : modules()) {
      powertoy.enable();
    }

    Trace::EventLaunch();

    result = run_message_loop();
  } catch (std::runtime_error err) {
    std::string err_what = err.what();
    MessageBox(NULL, std::wstring(err_what.begin(),err_what.end()).c_str(), TEXT("Error"), MB_OK | MB_ICONERROR);
    result = -1;
  }
  Trace::UnregisterProvider();
  return result;
}
