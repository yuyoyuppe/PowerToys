# Powertoy DLL Project For Visual Studio 2017

## Installation
Put the `Powertoy Module.zip` file inside `%%USERPROFILE%\Documents\Visual Studio 2017\Templates\ProjectTemplate\Visual C++` folder. The template will be available under `Visual C++` tab. Add the project to `src\modules\` folder for all the relative paths to work.

For the Powertoy to be loaded, it has to be added to the `known_dlls` map in `/src/runner/main.cpp`.

## Creating the template
Follow the ["How to: Create project templates"](https://docs.microsoft.com/en-us/visualstudio/ide/how-to-create-project-templates?view=vs-2017) guide from MSDN, using the `example_powertoy` project inside `powertoys/modules` filter as the base project. For icon select the `src/runner/svgs/icon.ico` file.
