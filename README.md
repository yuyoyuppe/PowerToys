# PowerToys

Two toys improving Windows experience:
  * When WinKey is held for more than 300ms an overlay is displayed showing various keyboard shortcuts.
  * Hovering with the mouse over maximize button displays popup window with extra options: for starters "maximize to new desktop".

## Build instructions
  * Use Visual Studio 2017, with Visual C++ features and `Windows 10 SDK (10.0.17763.0)` installed.
  * Open `powertoys.sln` in Visual Studio and build the `powertoys` project. PowerToys require admin privileges to run. In order to run and debug it from Visual Studio, it has to be started with elevated privileges.
  * For debugging outside of Visual Studio, copy the `svgs` folder from `src/runner/svgs` to the same path as `powertoys.exe`.

### Building the installer
  * Detailed instructions in the [PowerToysSetup project README.](installer/PowerToysSetup/README.md)

## Creating new PowerToys

To create a new PowerToy:

  * Install the [PowerToy Module project template](tools/project_template)
  * Add a new `PowerToy Module` project under `src/modules`,
  * Take a look at [the interface](src/modules/interface/powertoy_module_interface.h) and
    [the example PowerToy implementation](src/modules/example_powertoy/dllmain.cpp),
  * Each PowerToy is built as a DLL and in order to be loaded at run-time, the PowerToy's dll name needs to be added to the know_dlls map in the [src/runner/main.cpp](src/runner/main.cpp).
