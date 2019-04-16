#include "pch.h"
#include "functionalities.h"
#include <ShellScalingApi.h>
#include "d2d_window.h"
#pragma comment(lib, "shcore.lib")
#pragma comment(lib, "windowsapp")

//For PathRemoveFileSpec and PathCombine. Evaluate need for it afterwards.
#include <Shlwapi.h>
#pragma comment(lib, "shlwapi.lib")

bool CreateChildProcesses() {
  LPSTR child64=(LPSTR)"powertoys_hook_loader.exe";
  LPSTR child32=(LPSTR)"powertoys_hook_loader_32.exe";

  //TODO: Review way to get the path to the helper executables. Some buffer overflows possible here.
  HMODULE hModule = GetModuleHandle(NULL);
  CCHAR executablePath[MAX_PATH];
  GetModuleFileName(hModule, executablePath, MAX_PATH);
  PathRemoveFileSpec(executablePath);
  CCHAR executableHelper64Path[MAX_PATH];
  PathCombine(executableHelper64Path, executablePath, child64);
  CCHAR executableHelper32Path[MAX_PATH];
  PathCombine(executableHelper32Path, executablePath, child32);

  // Configures a job so that Windows kill all child processes when powertoys terminates.
  HANDLE hJobObj = CreateJobObject(NULL, "PowerToys Job");
  // TODO: This handle seems to be cleaned when the Application stops running, but check for further cleanup instructions.
  JOBOBJECT_EXTENDED_LIMIT_INFORMATION jobObjInfo;
  AssignProcessToJobObject(hJobObj, GetCurrentProcess());
  QueryInformationJobObject(hJobObj, JobObjectExtendedLimitInformation, &jobObjInfo, sizeof(jobObjInfo), NULL);
  jobObjInfo.BasicLimitInformation.LimitFlags |= JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;
  SetInformationJobObject(hJobObj, JobObjectExtendedLimitInformation, &jobObjInfo, sizeof(jobObjInfo));

  STARTUPINFO si64;
  PROCESS_INFORMATION pi64;
  ZeroMemory(&si64, sizeof(si64));
  si64.cb = sizeof(si64);
  ZeroMemory(&pi64, sizeof(pi64));

  STARTUPINFO si32;
  PROCESS_INFORMATION pi32;
  ZeroMemory(&si32, sizeof(si32));
  si32.cb = sizeof(si32);
  ZeroMemory(&pi32, sizeof(pi32));
  if (!CreateProcess(executableHelper64Path, executableHelper64Path, NULL, NULL, FALSE, 0, NULL, NULL, &si64, &pi64)) {
    ShowLastErrorMessage((LPTSTR)"CreateChildProcesses64", GetLastError());
    return false;
  }
  if (!CreateProcess(executableHelper32Path, executableHelper32Path, NULL, NULL, FALSE, 0, NULL, NULL, &si32, &pi32)) {
    ShowLastErrorMessage((LPTSTR)"CreateChildProcesses32", GetLastError());
    return false;
  }
  //TODO: Do proper cleanup of the handles before returning.
  return true;
}


int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
  winrt::init_apartment();
  winrt::check_hresult(SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE));
  // We will handle scaling ourselfs.
  HHOOK handle = NULL;
  try {
    start_winkey_handler();
    //start_mouse_handler();
    //CreateChildProcesses();
    int result = run_message_loop();
  return result;
  } catch (std::runtime_error err) {
    MessageBox(NULL, err.what(), "Error", MB_OK | MB_ICONERROR);
    return -1;
  }
}
