# Introduction

::: warning
Vue Excel is still in active development and is not production ready. Proceed with caution.
If you're interested in using it, please drop me a note [here](https://github.com/demsullivan/vue-excel/issues/1)
:::

Vue Excel brings the modern, reactive, component-based development of Vue.js to the world of Add-ins for Excel. Features include:

- Simplified management of Excel RequestContext
- Component-based architecture for working with Excel objects
- Use props and events for interacting with Excel
- Ability to drop down to the Office.js API if necessary
- Supports simple and complex routing cases

---

As a simple example, if you wanted to build an Add-in that reacts to data being changed in a Worksheet named "Sheet1", your code might look like this:

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

[Get started](./guide/) or [check out a working example](https://github.com/demsullivan/vue-excel-example).
