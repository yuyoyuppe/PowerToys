# Web project for the Settings UI

The UI for the PowerToys Settings is created using FabricUI.

## Build Commands

Here are the commands to build and test this project:

### To start the development server

```
npm install
npm run start
```

### Building and integrating into PowerToys settings project

```
npm run build -- --min --production
```

Copy the resulting `./dist` to `../editor/settings-html/dist`.
Also copy `index.html` to `../editor/settings-html/index.html` if you change it.

## Updating the icons

Icons inside `./src/icons/` were generated from the [Office UI Fabric Icons subset generation tool.](https://uifabricicons.azurewebsites.net/)

In case the subset needs to be changed, additional steps are needed to include the icon font in the built `dist/bundle.js`:
- Copy the inline font data taken from `src/icons/css/fabric-icons-inline.css` and place it in the `fontFace` `src` value in `src/icons/src/fabric-icons.ts`.

A list of the current icons in the subset can be seen in the `icons` object in `src/icons/src/fabric-icons.ts`.
