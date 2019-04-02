// main.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "VirtualDesktopsManagerInternalTest.h"
#include "VirtualDesktopsManagerTest.h"

int main()
{
    ::CoInitialize(NULL);

    VirtualDesktops::Tests::VirtualDesktopsManagerInternalTest testInternal;

    testInternal.EnumVirtualDesktops();
    testInternal.EnumAdjacentDesktops();
    testInternal.CurrentVirtualDesktop();
    testInternal.ManageVirtualDesktops(2000);

    VirtualDesktops::Tests::VirtualDesktopsManagerTest testPublic;

    testPublic.GetWindowDesktopId(GetConsoleWindow());

    return 0;
}

