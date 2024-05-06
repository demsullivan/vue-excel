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

## Quickstart

Get started by reading the documentation, or checking out the [example repo](https://github.com/demsullivan/vue-excel-example).
