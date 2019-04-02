#pragma once

#include "VirtualDesktopsAPI.h"

namespace VirtualDesktops
{
    namespace Tests
    {
        class VirtualDesktopsTestBase
        {
        protected:
            void printGuid(const GUID &guid);

            IServiceProvider* GetServiceProvider();
        };
    }
}