# Maximize to new desktop

# Introduction
Maximizes a Window to a new virtual desktop.

# Usage
When the mouse pointer hovers over a window's maximize/restore button, it shows a popup that creates a new virtual desktop and maximizes the window to the new virtual desktop.
If the Window is not in the primary virtual desktop, pressing clicking the popup returns the window to the primary virtual desktop and restores its original position.

# Options
These configurations can be edited from the PowerToys Settings screen:
- "How long to wait hovering the maximize button before showing the popup (ms)" - How many milliseconds to hover over a maximize/restore button before the MTND popup is shown.
- "Remove a virtual desktop when the last window is restored" - Whether to delete a virtual desktop after the last window in that desktop is restored to the primary desktop through the MTND popup.

# Known issues
MTND might not work correctly with windows that are running with elevated priveleges when PowerToys is not running with elevated priveleges.

# Code organization

#### [`dllmain.cpp`](./dllmain.cpp)
Contains DLL boilerplate code and implementation of the [PowerToys interface](/src/modules/interface/).

#### [`target_state.cpp`](./target_state.cpp)
Handles the events to transition between the overlay states.

#### [`virtual_desktops.cpp`](./virtual_desktops.cpp)
Contains the code for creating and moving a window between virtual desktops.

#### [`mouse_watcher.cpp`](./mouse_watcher.cpp)
Contains the code to detect if the maximize button is beneath the mouse pointer.

#### [`mouse_track_events.cpp`](./mouse_track_events.cpp)
Contains helper code to detect hover events in order to show the tooltip on the MTND popup.

#### [`trace.cpp`](./trace.cpp)
Contains code for telemetry.
