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
  static ComInit init;
}

