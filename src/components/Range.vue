<script setup lang="ts">
import { inject, onMounted, shallowRef, watch } from 'vue'
import type Context from 'vue-excel/Context'

// REFS AND PROPS
type Props = {
  address?: string
  name?: string
  value: string[][] | null | undefined
}

const props = defineProps<Props>()

const context: Context = inject('vueExcel.context') as Context
const binding = shallowRef<Excel.Binding>()

watch(
  () => props.value,
  async (value, oldValue) => {
    // TODO: detect the size of the range, and handle values accordingly.
    // For example, if it's a single cell, use single values.
    // If it's a column, use an array of values.
    // If it's a row, also use an array of values.
    // If it's two dimensional, use an array of arrays.
    if (oldValue == undefined && (value == null || value == undefined)) return
    if (!binding.value) return

    const range = (binding.value as Excel.Binding).getRange()

    if (oldValue != null && oldValue != undefined && (value == null || value == undefined)) {
      range.clear()
    } else if (value instanceof Array) {
      range.values = value
    }

    await context.sync()
  }
)

// EMITS AND EVENTS
type Emits = {
  dataChanged: [event: Excel.BindingDataChangedEventArgs]
  selectionChanged: [event: Excel.BindingSelectionChangedEventArgs]
}

const emit = defineEmits<Emits>()

// FUNCTIONS
onMounted(async () => {
  const { xlBinding } = await context.fetch(async (ctx: Excel.RequestContext) => {
    let xlBinding

    if (props.address) {
      xlBinding = ctx.workbook.bindings.add(props.address, Excel.BindingType.range, props.address)
    } else if (props.name) {
      xlBinding = ctx.workbook.bindings.addFromNamedItem(props.name, Excel.BindingType.range, props.name)
    } else {
      console.error('You must pass either address or name props to a Range component!')
      return { xlBinding: undefined }
    }

    xlBinding.onDataChanged.add(async (event: Excel.BindingDataChangedEventArgs) => {
      emit('dataChanged', event)
    })

    xlBinding.onSelectionChanged.add(async (event: Excel.BindingSelectionChangedEventArgs) => {
      emit('selectionChanged', event)
    })

    return { xlBinding }
  })

  binding.value = xlBinding
})
</script>

<template></template>
