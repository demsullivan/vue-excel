# Installation

The following instructions were adapted from [this tutorial](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-vue) from Microsoft.

Currently, installation is very manual. Hopefully in the near future, there will be a generator that automatically sets all of this up for you.

## Install dependencies

```bash
npm install -g yo generator-office
```

## Create a new Vue app

```bash
npm create vue@latest
```

## Generate the Office manifest file

Enter your new project directory, and run:

```bash
yo office
```

Select the following options from the Office.js generator prompts:

- Choose a project type: `Office Add-in project containing the manifest only`
- What do you want to name your add-in? `addin`
- Which Office client application would you like to support? `Excel`

Then, copy the `manifest.xml` file and `assets` directory to the proper locations.

```bash
mv addin/manifest.xml .
mv addin/assets public
rm -rf addin
```

## Install vue-excel and other Office dependencies

```bash
npm install --save demsullivan/vue-excel#main office-toolbox
npm install -D @types/office-js @types/office-runtime
```

## Set up HTTPS

Microsoft recommends using HTTPS for Office Add-ins. For development and testing, there is a tool for
generating and using self-signed certificates that also configures Office to trust these certificates.
See [this section](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-vue?view=excel-js-preview#secure-the-app)
of the Microsoft docs for more details.

1. Update your `vite.config.ts` as follows:

   ::: code-group

   ```typescript [vite.config.ts]
   import { fileURLToPath, URL } from 'node:url'
   import fs from 'fs'
   import path from 'path'
   import { homedir } from 'os'
   import { defineConfig } from 'vite'
   import vue from '@vitejs/plugin-vue'

   // https://vitejs.dev/config/
   export default defineConfig({
     plugins: [vue()],
     resolve: {
       alias: {
         '@': fileURLToPath(new URL('./src', import.meta.url))
       }
     },
     server: {
       port: 3000,
       https: {
         key: fs.readFileSync(path.resolve(`${homedir()}/.office-addin-dev-certs/localhost.key`)),
         cert: fs.readFileSync(path.resolve(`${homedir()}/.office-addin-dev-certs/localhost.crt`)),
         ca: fs.readFileSync(path.resolve(`${homedir()}/.office-addin-dev-certs/ca.crt`))
       }
     }
   })
   ```

   :::

2. Run the dev certificate generator:

   ```bash
   npx office-addin-dev-certs install
   ```

## Update the code

First, update your `manifest.xml` so that it uses the correct URL for the Taskpane. Find the `<bt:Urls>` tag inside the `<Resources>` tag. Locate the tag with the id `Taskpane.Url` and update the `DefaultValue` attribute to be `https://localhost:3000/index.html`.

For example:

::: code-group

```xml [manifest.xml]
<bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html" /> // [!code --]
<bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/index.html" /> // [!code ++]
```

:::

Lastly, replace the following files with the contents below.

::: code-group

```typescript [env.d.ts]
/// <reference types="vite/client" />
/// <reference types="office-js" /> // [!code ++]
```

```html [index.html]
<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <link rel="icon" href="/favicon.ico" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Vite App</title>
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script> // [!code ++]
  </head>
  <body>
    <div id="app"></div>
    <script type="module" src="/src/main.ts"></script>
  </body>
</html>
```

```typescript [src/main.ts]
import './assets/main.css'

import { createApp } from 'vue'
import App from './App.vue'
import { connectExcel } from 'vue-excel' // [!code ++]

createApp(App).mount('#app') // [!code --]
window.Office.onReady(async () => {
  // [!code ++]
  const vueExcel = await connectExcel() // [!code ++]
  createApp(App).use(vueExcel, {}).mount('#app') // [!code ++]
}) // [!code ++]
```

:::
