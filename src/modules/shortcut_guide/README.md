# Windows Key Shortcut Guide

# Introduction
The Windows Key Shortcut Guide shows common keyboard shortcuts that use the Windows key.

# Usage
Press and hold the keyboard Windows key for about 1 second, an overlay appears showing keyboard shortcuts that use the Windows Key:
- Shortcuts for changing the position of the active window.
- Common Windows shortcuts.
- Taskbar shortcuts.

Releasing the Windows key will make the overlay disappear.

![Image of the Overlay](/doc/images/shortcut_guide/usage.png)

The keyboard shortcuts can be used while the guide is being shown.

# Options
These configurations can be edited from the PowerToys Settings screen:
- "How long to press the Windows key before showing the Shortcut Guide (ms)" - How many milliseconds to press the Windows key before the Shortcut Guide is shown.
- "Opacity of the Shortcut Guide's overlay background (%)" - Changing this setting controls the opacity of the Shortcut Guide's overlay background, occluding the work environment beneath the Shortcut Guide to different degrees.

![Image of the Options](/doc/images/shortcut_guide/settings.png)

# Known issues
The Shortcut Guide hasn't been localized yet. Some of the shortcuts shown may not apply to localized Windows versions.

# Code organization

#### [`dllmain.cpp`](./dllmain.cpp)
Contains DLL boilerplate code.

#### [`shortcut_guide.cpp`](./shortcut_guide.cpp)
Contains the main PowerToys module interface code.

#### [`overlay_window.cpp`](./overlay_window.cpp)
Contains the code for creating the overlay window.

#### [`keyboard_state.cpp`](./keyboard_state.cpp)
Contains helper methods for checking the current state of the keyboard.

#### [`target_state.cpp`](./target_state.cpp)
Handles the events to transition between the overlay states.

#### [`trace.cpp`](./trace.cpp)
Contains code for telemetry.
