# vue-excel

A Vue plugin for building declarative, reactive Office.js Excel Add-ins.

> [!CAUTION]
> This library has not been released yet and is still under active development.
> If you're interested in using it, please drop me a note [here](https://github.com/demsullivan/excelsius/issues/1)

## Install

Install with your favourite package manager:

```
npm install --save vue-excel
```

## Getting Started

Spin up an Office.js Excel Add-in using Vue, following the instructions [here](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-vue)

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

## Working with Ranges

## Workbook Props

## Reusing Components with Multiple Worksheets
