# vue-excel

A Vue plugin for building declarative, reactive Office.js Excel Add-ins.

> [!CAUTION]
> This library has not been released yet and is still under active development.
> If you're interested in using it, please drop me a note [here](https://github.com/demsullivan/vue-excel/issues/1)

## What is it?

Microsoft allows developers to build Add-ins for Office using web technologies such as Javascript, HTML, and CSS.
I found the Office.js API to be a bit tedious to work with, and wanted a simpler and more modern interface for
interacting with Excel and building my own Add-in for Excel.

Thus, vue-excel was born. Vue-excel allows you to build modern, reactive Add-ins for Excel, using Vue's familiar
components, props and events patterns.

As an example, if you wanted to build an Add-in that reacts to data being changed in a Worksheet, your code
might look like this:

```vue
<script setup lang="ts">
function sheetChanged(event: Excel.WorksheetChangedEventArgs) {
  const newValue = event.details.valueAfter
  // do something with the new value
}
</script>

<template>
  <Workbook>
    <Worksheet name="Sheet1" @changed="sheetChanged" />
  </Workbook>
</template>
```

## Working Example

Check out the [vue-excel-example](https://github.com/demsullivan/vue-excel-example) repo for a working examples. That
repo was created by following the Getting Started documentation below, with some additional functionality examples included.

## Getting Started

The following instructions were adapted from [this tutorial](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-vue) from Microsoft.

### Install the necessary dependencies

```
npm install -g yo generator-office
```

### Generate a new Vue app

In this example, we'll name the project `vue-excel-example`.

```
npm create vue@latest
```

### Generate the Office manifest file

This allows Office to load your code as an Add-in.

```
cd vue-excel-example
yo office
```

When prompted, you'll need to choose a different directory than where you created your new Vue app.

Select the following options from the Office.js generator prompts:
- Choose a project type: `Office Add-in project containing the manifest only`
- What do you want to name your add-in? `add-in`
- Which Office client application would you like to support? `Excel`

### Copy the `manifest.xml` and `assets` directory to your project root

```
mv add-in/manifest.xml .
mv add-in/assets public
rm -rf add-in
```

### Install vue-excel and other office dependencies

```
npm install --save demsullivan/vue-excel#main office-toolbox
npm install -D @types/office-js @types/office-runtime
```

### Set up HTTPS

Microsoft recommends using HTTPS for Office Add-ins. For development and testing, there is a tool for 
generating and using self-signed certificates that also configures Office to trust these certificates. 
See [this section](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-vue?view=excel-js-preview#secure-the-app) 
of the Microsoft docs for more details.

1. Update your `vite.config.ts` as follows:

    ```typescript
    import { fileURLToPath, URL } from 'node:url'
    import fs from 'fs'
    import path from 'path'
    import { homedir } from 'os'
    import { defineConfig } from 'vite'
    import vue from '@vitejs/plugin-vue'

    // https://vitejs.dev/config/
    export default defineConfig({
      plugins: [
        vue(),
      ],
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

2. Run the dev certificate generator:

    ```
    npx office-addin-dev-certs install
    ```

### Update the app

1. Add a reference to office-js types to `env.d.ts`

    ```typescript
    /// <reference types="office-js" />
    ```

2. Open `index.html` and add the following `<script>` tag immediately before the `</head>` tag

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

3. Open `manifest.xml` and find the `<bt:Urls>` tags inside the `<Resources>` tag. Locate the
   tag with id `Taskpane.Url` and update the `DefaultValue` attribute to be
   `https://localhost:3000/index.html`

    ```html
    <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/index.html" />
    ```

4. Open `src/main.ts` and replace the contents with the following code

    ```typescript
    import './assets/main.css'

    import { createApp } from 'vue'
    import App from './App.vue'
    import { connectExcel } from 'vue-excel'

    window.Office.onReady(async () => {
      const vueExcel = await connectExcel()

      createApp(App)
      .use(vueExcel, {})
      .mount('#app')
    })

    ```

5. Open `src/App.vue` and replace the contents with the following code

    ```vue
    <script setup lang="ts">
    import { ref } from 'vue'

    const rangeValue = ref()

    function sheetChanged(event: Excel.WorksheetChangedEventArgs) {
      rangeValue.value = [[event.details.valueAfter]]
    }

    function updateRange() {
      rangeValue.value = [['Button clicked!']]
    }
    </script>

    <template>
      <button @click="updateRange">Click me!</button>
      <Workbook>
        <Worksheet name="Sheet1" @changed="sheetChanged">
          <Range address="A1" :value="rangeValue" />
        </Worksheet>
      </Workbook>
    </template>
    ```

## Using your Add-In in Excel

### Run the dev server

```
npm run dev
```

### Sideload your Add-in within Excel

Follow the instructions from Microsoft on sideloading an Add-in here:
https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins


## Rendering a Task Pane

You can render content in Excel's Taskpane by simply adding UI components to the root `App` component.
For example:

```vue
<script setup lang="ts">
// src/App.vue
import MyComponent from './components/MyComponent.vue'
</script>

<template>
  <Workbook>
    <MyComponent />
    <Worksheet name="Sheet1">
  </Workbook>
</template>
```

### Making the Taskpane Interactive

## Working with Ranges

## Workbook Props

## Reusing Components with Multiple Worksheets
