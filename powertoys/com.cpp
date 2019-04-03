#include "com.h"
#include <Windows.h>
#include <stdexcept>
struct ComInit {
  ComInit() {
    if (!SUCCEEDED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED))) {
      throw std::runtime_error("CoInitializeEx failed");
    }
  }
  ~ComInit() {
    CoUninitialize();
  }
};

void init_com() {
  // C++ guarantees this will be initialized once in a thread-safe way
  static ComInit init;
}

