#include "pch.h"
#include "debug_trace.h"
#include <iostream>
#include <fstream>
#include <filesystem>

using namespace std;
using namespace filesystem;

#ifdef _DEBUG
#define LOGDIR "C:\\PowerToysTrace"
#define LOGPATHRUNNER "C:\\PowerToysTrace\\runner.txt"
#define LOGPATHSETTINGS "C:\\PowerToysTrace\\settings.txt"
FILE* file_log;
#endif

void init_debug_trace_runner() {
#ifdef _DEBUG
  if (!exists(LOGDIR)) {
    create_directory(LOGDIR);
  }
  file_log = _fsopen(LOGPATHRUNNER, "w", _SH_DENYWR);
#endif
}

void init_debug_trace_settings() {
#ifdef _DEBUG
  if (!exists(LOGDIR)) {
    create_directory(LOGDIR);
  }
  file_log = _fsopen(LOGPATHSETTINGS, "w", _SH_DENYWR);
#endif
}

void close_debug_trace() {
#ifdef _DEBUG
  fclose(file_log);
#endif
}

#ifdef _DEBUG
void _trace(const char* function, const char* file, const int line, const char* message) {
  if (file_log != nullptr) {
    fprintf(file_log, "%s [line %ld] %s: %s\n", file, line, function, message);
    fflush(file_log);
  }
}

void _trace_last_error(const char* function, const char* file, const int line, const char* message) {
  string msg = message;
  DWORD err_code = GetLastError();
  msg += " - err: " + to_string(err_code);

  LPVOID buffer;
  if (FormatMessageA(
      FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
      NULL,
      err_code,
      MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
      (LPSTR)&buffer,
      0,
      NULL) > 0) {
    msg += " ";
    msg += (char*)buffer;
    LocalFree(buffer);
  }
  _trace(function, file, line, msg.c_str());
}

void _trace_bool(const char* function, const char* file, const int line, const char* message, bool value) {
  string msg = message;
  msg += value ? " true" : " false";
  _trace(function, file, line, msg.c_str());
}

void _trace_ptr(const char* function, const char* file, const int line, const char* message, void* value) {
  string msg = message;
  msg += value ? " valid ptr" : " nullptr";
  _trace(function, file, line, msg.c_str());
}

#endif
