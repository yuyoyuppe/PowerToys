#include "pch.h"
#include "tray_icon.h"
#include <ShellScalingApi.h>

#include "powertoy_module.h"
#include "lowlevel_keyboard_event.h"
#include <filesystem>

TRACELOGGING_DEFINE_PROVIDER(
    g_hProvider,
    "Microsoft.PowerToys",
    // {38e8889b-9731-53f5-e901-e8a7c1753074}
    (0x38e8889b, 0x9731, 0x53f5, 0xe9, 0x01, 0xe8, 0xa7, 0xc1, 0x75, 0x30, 0x74),
    TraceLoggingOptionProjectTelemetry());

#if _DEBUG && _WIN64
#include "unhandled_exception_handler.h"
#endif

void chdir_current_executable() {
  // Change current directory to the path of the executable.
  CCHAR executable_path[MAX_PATH];
  GetModuleFileName(NULL, executable_path, MAX_PATH);
  PathRemoveFileSpec(executable_path);
  if(!SetCurrentDirectory(executable_path)) {
    show_last_error_message((LPSTR)"Change Directory to Executable Path", GetLastError());
  }
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
  #if _DEBUG && _WIN64
  //Global error handlers to diagnose errors.
  //We prefer this not not show any longer until there's a bug to diagnose.
  //init_global_error_handlers();
  #endif
  TraceLoggingRegister(g_hProvider);
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
      L"move_to_new_desktop.dll" 
    };
    for (auto& file : std::filesystem::directory_iterator("modules/")) {
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

    TraceLoggingWrite(
        g_hProvider,
        "PowerToysLaunch",
        TraceLoggingDescription("Successful launch of PowerToys"),
        ProjectTelemetryPrivacyDataTag(ProjectTelemetryTag_ProductAndServicePerformance),
        TraceLoggingBoolean(TRUE, "UTCReplace_AppSessionGuid"),
        TraceLoggingKeyword(PROJECT_KEYWORD_MEASURE));

    result = run_message_loop();
  } catch (std::runtime_error err) {
    MessageBox(NULL, err.what(), "Error", MB_OK | MB_ICONERROR);
    result = -1;
  }
  TraceLoggingUnregister(g_hProvider);
  return result;
}
