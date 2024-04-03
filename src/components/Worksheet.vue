<script setup lang="ts">
import { onMounted, inject, ref, type Ref, watch, type EmitsOptions, computed } from 'vue'
import VueExcel from '../VueExcel'

// REFS AND PROPS
const vueExcel: VueExcel = inject('vueExcel') as VueExcel
const worksheet = ref<Excel.Worksheet>()

type Props = {
  name: string
}

const props = defineProps<Props>()

// EMITS AND EVENTS
type Emits = {
  changed: [event: any]
}

type WorksheetEvent = "onChanged"
type WorksheetEventArgs = Excel.WorksheetChangedEventArgs

const emitEvents: Record<keyof Emits, WorksheetEvent> = {
  changed: 'onChanged'
}

const emit = defineEmits<Emits>()

// FUNCTIONS
async function emitEvent(emitName: keyof Emits, event: WorksheetEventArgs) {
  emit(emitName, event)
}

function setupEventListeners() {
  const sheet = worksheet.value
  if (sheet == null) return

  const emitNames = Object.keys(emitEvents) as (keyof Emits)[]

  emitNames.forEach((emitName: keyof Emits) => {
    const eventName: WorksheetEvent = emitEvents[emitName]
    sheet[eventName].add(emitEvent.bind({}, emitName))
  })
}

onMounted(async() => {
  return await vueExcel.excel.run(async (ctx: Excel.RequestContext) => {
    const excelWorksheet = ctx.workbook.worksheets.getItem(props.name).load()
    await ctx.sync()
    worksheet.value = excelWorksheet
    setupEventListeners()
  })
})
</script>

<template>
  <slot></slot>
</template>