# Powertoy DLL Project For Visual Studio 2017

## Installation
Put the `Powertoy Module.zip` file inside `%%USERPROFILE%\Documents\Visual Studio 2017\Templates\ProjectTemplate\Visual C++` folder. The template will be available under `Visual C++` tab. Add the project to `src\modules\` folder for all the relative paths to work.

For the Powertoy to be loaded, it has to be added to the `known_dlls` map in [/src/runner/main.cpp](../src/runner/main.cpp).

