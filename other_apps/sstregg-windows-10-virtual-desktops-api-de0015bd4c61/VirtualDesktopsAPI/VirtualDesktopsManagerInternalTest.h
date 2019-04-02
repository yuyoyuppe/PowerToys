#pragma once

#include "VirtualDesktopsTestBase.h"

namespace VirtualDesktops
{
    namespace Tests
    {
        class VirtualDesktopsManagerInternalTest final : public VirtualDesktopsTestBase
        {
        private:
            IServiceProvider *m_pServiceProvider;
            API::IVirtualDesktopManagerInternal *m_pDesktopManager;

        public:
            VirtualDesktopsManagerInternalTest();

            ~VirtualDesktopsManagerInternalTest();

            void EnumVirtualDesktops();

            void CurrentVirtualDesktop();

            void EnumAdjacentDesktops();

            void ManageVirtualDesktops(DWORD dwSleepMiliseconds);

            void Do();
        };
    }
}
