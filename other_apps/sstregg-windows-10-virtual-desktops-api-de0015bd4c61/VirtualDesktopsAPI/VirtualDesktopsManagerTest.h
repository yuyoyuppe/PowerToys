#pragma once

#include "VirtualDesktopsTestBase.h"

namespace VirtualDesktops
{
    namespace Tests
    {
        class VirtualDesktopsManagerTest final : public VirtualDesktopsTestBase
        {
        private:
            IServiceProvider *m_pServiceProvider;
            API::IVirtualDesktopManager *m_pDesktopManager;

        public:
            VirtualDesktopsManagerTest();

            ~VirtualDesktopsManagerTest();

            void GetWindowDesktopId(HWND hWnd);
        };
    }
}