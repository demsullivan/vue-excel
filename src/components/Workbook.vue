<script setup lang="ts">
import { inject } from 'vue';
import type VueExcel from '../VueExcel'

const vueExcel: VueExcel = inject('vueExcel') as VueExcel

/*
Static Routing Options

- Worksheet binds to sheet by name
- Worksheet only renders its children if its the active worksheet
- Useful for statically named worksheets
- MyComponent could access worksheet via injection
<Workbook>
  <Worksheet name="Sheet1">
    <MyComponent />
  </Worksheet>
</Workbook>

- MyComponent defines Worksheet
- Puts routing/display logic under control of MyComponent
- Allows MyComponent to access Worksheet props/events directly
<Workbook>
  <MyComponent />
</Workbook>

Dynamic Routing Options
<Workbook>
  <Worksheet :has-name="{ account: 'thing' }">
    <ThingComponent />
  </Worksheet>
</Workbook>

*/
</script>

<template>
  <slot v-if="vueExcel"></slot>
</template>

