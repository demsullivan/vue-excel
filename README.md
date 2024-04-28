# vue-excel

A Vue plugin for building declarative, reactive Office.js Excel Add-ins.

> [!CAUTION]
> This library has not been released yet and is still under active development.
> If you're interested in using it, please drop me a note [here](https://github.com/demsullivan/vue-excel/issues/1)

## Install

Install with your favourite package manager:

```
npm install --save vue-excel
```

## Getting Started

The following instructions were adapted from [this tutorial](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-vue) from Microsoft.


1. Install the necessary dependencies

```
npm install -g yo generator-office
```

2. Generate a new Vue app

```
npm create vue@latest
```

3. Generate the Office manifest file, which allows Office to load your code as an Add-in

```
yo office
```

Select the following options from the Office.js generator prompts:
- Choose a project type: `Office Add-in project containing the manifest only`
- What do you want to name your add-in? `vue-excel-examples`
- Which Office client application would you like to support? `Excel`


4. Set up HTTPS

Microsoft recommends using HTTPS for Office Add-ins. For development and testing, there is a tool for generating and using self-signed certificates that also configures Office to trust these certificates. See [this section](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-vue?view=excel-js-preview#secure-the-app) of the Microsoft docs for more details.

- 


Update your initialization code to install vue-excel by editing `main.ts` - for example:

```typescript
import { createApp } from 'vue'
import App from './App.vue'
import { connectExcel } from 'vue-excel'

window.Office.onReady(async () => {
  const vueExcel = connectExcel()

  createApp(App)
    .use(vueExcel, {
      components: []
    })
    .mount("#app")
})
```
Lastly, update your `App` root component in `App.vue` to use components from vue-excel:

```vue
<script setup lang="ts">

function sheetChanged(event: Excel.WorksheetChangedEventArgs) {
  console.log("sheet changed!")
  console.dir(sheet)
}

</script>

<template>
  <Workbook>
    <Worksheet name="Sheet1" @changed="sheetChanged" />
  </Workbook>
</template>
```

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
