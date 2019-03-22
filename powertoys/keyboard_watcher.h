#pragma once

#include<functional>

void start_winkey_watcher(int ms_delay, std::function<void()> on_held, std::function<void()> on_released);
