<script setup lang="ts">
import { onMounted, inject, shallowRef, provide } from 'vue'
import type Context from '@/Context'
import { EmitFlags } from 'typescript'

// REFS AND PROPS
const context: Context = inject('vueExcel.context') as Context
const worksheet = shallowRef<Excel.Worksheet>()

provide('vueExcel.scope.worksheet', worksheet)
defineExpose({ worksheet })

type Props = {
  name: string
}

const props = defineProps<Props>()

// EMITS AND EVENTS

type EmitEventMap = {
  changed: { handler: 'onChanged'; eventArgs: Excel.WorksheetChangedEventArgs }
  selectionChanged: {
    handler: 'onSelectionChanged'
    eventArgs: Excel.WorksheetSelectionChangedEventArgs
  }
}

type Emits = {
  changed: [event: Excel.WorksheetChangedEventArgs]
  selectionChanged: [event: Excel.WorksheetSelectionChangedEventArgs]
}

type WorksheetEventHandler = {
  [P in keyof EmitEventMap]: EmitEventMap[P]['handler']
}[keyof EmitEventMap]

type WorksheetEventArgs = {
  [P in keyof EmitEventMap]: EmitEventMap[P]['eventArgs']
}[keyof EmitEventMap]

const emitEvents: Record<keyof Emits, WorksheetEventHandler> = {
  changed: 'onChanged',
  selectionChanged: 'onSelectionChanged'
}

const emit = defineEmits<Emits>()

// FUNCTIONS
async function emitEvent(emitName: keyof Emits, event: WorksheetEventArgs) {
  // @ts-ignore - TS doesn't like the dynamic emit name because of how defineEmits is typed.
  emit(emitName, event)
}

function setupEventListeners() {
  const sheet = worksheet.value
  if (sheet == null) return

  const emitNames = Object.keys(emitEvents) as (keyof Emits)[]

  emitNames.forEach((emitName: keyof Emits) => {
    const eventName: WorksheetEventHandler = emitEvents[emitName]
    sheet[eventName].add(emitEvent.bind({}, emitName))
  })
}

onMounted(async () => {
  const { xlWorksheet } = await context.fetch(async (ctx) => ({
    xlWorksheet: ctx.workbook.worksheets.getItem(props.name)
  }))

  worksheet.value = xlWorksheet
  setupEventListeners()
})
</script>

<template>
  <slot></slot>
</template>
