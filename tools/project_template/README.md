# PowerToy DLL Project For Visual Studio 2017

## Installation

- Put the `PowerToy Module.zip` file inside the `%USERPROFILE%\Documents\Visual Studio 2017\Templates\ProjectTemplate\Visual C++` folder. 
- The template will be available in Visual Studio, when adding a new project, under the `Visual C++` tab.

## Create a new PowerToy Module

- Add the new PowerToy project to the `src\modules\` folder for all the relative paths to work.
- For the module interface implementation take a look at [the interface](../../src/modules/interface/powertoy_module_interface.h) and
    [the example PowerToy implementation](../../src/modules/example_powertoy/dllmain.cpp)
- Each PowerToy is built as a dll and in order to be loaded at run-time, the PowerToy's dll name needs to be added to the `know_dlls` map in [/src/runner/main.cpp](../../src/runner/main.cpp).
