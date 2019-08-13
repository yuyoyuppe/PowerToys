# PowerToy DLL Project For Visual Studio 2017

## Installation

- Put the `PowerToy Module.zip` file inside the `%USERPROFILE%\Documents\Visual Studio 2017\Templates\ProjectTemplate\Visual C++` folder. 
- The template will be available in Visual Studio, when adding a new project, under the `Visual C++` tab.

## Create a new PowerToy Module

- Add the new PowerToy project to the `src\modules\` folder for all the relative paths to work.
- For the module interface implementation take a look at [the interface](../../src/modules/interface/powertoy_module_interface.h) and
    [the example PowerToy implementation](../../src/modules/example_powertoy/dllmain.cpp)
- Each PowerToy is built as a DLL and in order to be loaded at run-time, the PowerToy's DLL name needs to be added to the `known_dlls` map in [/src/runner/main.cpp](../../src/runner/main.cpp).

## DPI Awareness

All PowerToy modules need to be DPI aware and calculate dimensions and positions of the UI elements using the Windows API for DPI awareness.
The src/common library has some helpers that you can use and extend:
 - dpi_aware.h, dpi_aware.cpp
 - monitors.h, monitors.cpp

## PowerToy settings

PowerToys provides a settings infrastructure to add a settings page for new modules.
The PowerToys Settings application is accessed from the PowerToys tray icon, it provides a global settings page and a dedicated settings page for each module.
The PowerToys settings API provides a way to define the required controls for the module's settings page and methods to read and persist the settings values. A module may need a more complex way to configure the user's preferences, in that case it can provide its own custom settings editor that can be invoked from the module's settings page through a dedicated button.

### Settings architecture overview
TODO

### Custom editor
TODO

### How to add your module's settings page
TODO


## Add a new PowerToy to the Installer

In the `installer` folder, open the `PowerToysSetup.sln` solution.
Under the `PowerToysSetup` project, edit `Product.wxs`.
You will need to add a component for your module DLL. Search for `Module_ShortcutGuide` to see where to add the component declaration and where to reference that declaration so the DLL is added to the installer.
Each component requires a newly generated GUID (you can use the Visual Studio integrated tool to generate one).
Repeat the process for each extra file your PowerToy module requires.
If your PowerToy comes with a subfolder containing for example images, follow the example of the `PowerToysSvgs` component.
