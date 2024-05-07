<script setup lang="ts">
import { onBeforeMount, onBeforeUnmount, inject, shallowRef, provide, watch, computed } from 'vue'
import type { VueExcelGlobalState } from '@vue-excel/index'

////////// TYPES //////////
type Props = {
  name: string
}

type Emits = {
  calculated: [event: Excel.WorksheetCalculatedEventArgs, handlerName: 'onCalculated']
  changed: [event: Excel.WorksheetChangedEventArgs, handlerName: 'onChanged']
  columnSorted: [event: Excel.WorksheetColumnSortedEventArgs, handlerName: 'onColumnSorted']
  formatChanged: [event: Excel.WorksheetFormatChangedEventArgs, handlerName: 'onFormatChanged']
  formulaChanged: [event: Excel.WorksheetFormulaChangedEventArgs, handlerName: 'onFormulaChanged']
  rowHiddenChanged: [event: Excel.WorksheetRowHiddenChangedEventArgs, handlerName: 'onRowHiddenChanged']
  rowSorted: [event: Excel.WorksheetRowSortedEventArgs, handlerName: 'onRowSorted']
  selectionChanged: [event: Excel.WorksheetSelectionChangedEventArgs, handlerName: 'onSelectionChanged']
  singleClicked: [event: Excel.WorksheetSingleClickedEventArgs, handlerName: 'onSingleClicked']
  visibilityChanged: [event: Excel.WorksheetVisibilityChangedEventArgs, handlerName: 'onVisibilityChanged']
}

type WorksheetEventHandler = {
  [P in keyof Emits]: Emits[P][1]
}[keyof Emits]

type WorksheetEventArgs = {
  [P in keyof Emits]: Emits[P][0]
}[keyof Emits]

////////// REFS AND PROPS //////////

const vueExcel = inject('vueExcel') as VueExcelGlobalState
const context = vueExcel.context
const activeWorksheet = vueExcel.activeWorksheet

const worksheet = shallowRef<Excel.Worksheet>()
const props = defineProps<Props>()
const isActive = computed<boolean>(() => {
  return props.name == activeWorksheet.value?.name
})

provide('vueExcel.scope.worksheet', worksheet)

////////// EMITS //////////
const emitEvents: Record<keyof Emits, WorksheetEventHandler> = {
  calculated: 'onCalculated',
  changed: 'onChanged',
  columnSorted: 'onColumnSorted',
  formatChanged: 'onFormatChanged',
  formulaChanged: 'onFormulaChanged',
  rowHiddenChanged: 'onRowHiddenChanged',
  rowSorted: 'onRowSorted',
  selectionChanged: 'onSelectionChanged',
  singleClicked: 'onSingleClicked',
  visibilityChanged: 'onVisibilityChanged'
}

const emit = defineEmits<Emits>()

async function emitEvent(emitName: keyof Emits, event: WorksheetEventArgs) {
  // @ts-ignore - TS doesn't like the dynamic emit name because of how defineEmits is typed.
  emit(emitName, event, emitEvents[emitName])
}

////////// LIFECYCLE HOOKS //////////
onBeforeMount(async () => {
  const { xlWorksheet } = await context.fetch(async (ctx) => ({
    xlWorksheet: ctx.workbook.worksheets.getItem(props.name)
  }))

  const emitNames = Object.keys(emitEvents) as (keyof Emits)[]

  emitNames.forEach((emitName: keyof Emits) => {
    const eventName: WorksheetEventHandler = emitEvents[emitName]
    xlWorksheet[eventName].add(emitEvent.bind({}, emitName))
  })

  await context.sync()

  worksheet.value = xlWorksheet
})

onBeforeUnmount(async () => {
  await context.run(async (ctx) => {
    const emitNames = Object.keys(emitEvents) as (keyof Emits)[]

    emitNames.forEach((emitName: keyof Emits) => {
      const eventName: WorksheetEventHandler = emitEvents[emitName]
      if (worksheet.value) {
        worksheet.value[eventName].remove(emitEvent.bind({}, emitName))
      }
    })
  })
})
</script>

<template>
  <slot></slot>
</template>
