#pragma once
#include "windef.h"

class DPIAware {
private:
  static const int DEFAULT_DPI = 96;

public:
  static void Convert(HMONITOR monitor_handle, int &width, int &high);
};
