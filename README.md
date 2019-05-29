# PowerToys

Two toys improving Windows experience:
  * When WinKey is held for more than 300ms an overlay is displayed showing various keyboard shortcuts.
  * Hovering with the mouse over maximize button displays popup window with extra options: for starters "maximize to new desktop".

## Build instructions
  * Use Visual Studio 2017, with Visual C++ features and `Windows 10 SDK (10.0.17763.0)` installed.
  * Open `powertoys.sln` in Visual Studio and build the `powertoys` project. Powertoys can be run from Visual Studio.
  * For debugging outside of Visual Studio, copy the `svgs` folder from `powertoys/svgs` to the same path as `powertoys.exe`.

### Building the installer
  * Detailed instructions in the [PowerToysSetup project README.](PowerToysSetup/README.md)
