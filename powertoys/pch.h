#pragma once
#include <winrt/base.h>
#include <Windows.h>
#include <Shobjidl.h>
#include <Shlwapi.h>
#include <string>
#include <algorithm>
#include <chrono>
#include <mutex>
#include <thread>
#include <functional>
#include <condition_variable>
#include <stdexcept>
#include <tuple>

#include "utils.h"