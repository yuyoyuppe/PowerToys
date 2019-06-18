// dllmain.cpp : Defines the entry point for the DLL application.
#include "pch.h"
#include <mutex>
#include "overlay_window.h"
#include "d2d_overlay_window.h"

BOOL APIENTRY DllMain(HMODULE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved) {
  switch (ul_reason_for_call) {
  case DLL_PROCESS_ATTACH:
  case DLL_THREAD_ATTACH:
  case DLL_THREAD_DETACH:
  case DLL_PROCESS_DETACH:
    break;
  }
  return TRUE;
}

extern "C" __declspec(dllexport) PowertoyModuleIface*  __cdecl powertoy_create() {
  if (!instance) {
    instance = new OverlayWindow();
    return instance;
  } else {
    return nullptr;
  }
}