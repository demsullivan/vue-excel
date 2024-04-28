<script setup lang="ts">
import { inject, onMounted, ref, shallowRef, watch } from 'vue'
import type VueExcel from '../VueExcel';


// REFS AND PROPS
type Props = {
  address?: string
  name?: string
  value: string[][] | null | undefined
}

const props = defineProps<Props>()

const vueExcel: VueExcel = inject('vueExcel') as VueExcel
const binding = shallowRef<Excel.Binding>()

watch(
  () => props.value,
  async (value, oldValue) => {
    if (oldValue == undefined && (value == null || value == undefined)) return
    if (!binding.value) return

    return await vueExcel.excel.run(async (ctx: Excel.RequestContext) => {
      const range = (binding.value as Excel.Binding).getRange()

      if (oldValue != null && oldValue != undefined && (value == null || value == undefined)) {
        range.clear()
      } else if (value instanceof Array) {
        range.values = value
      }

      await ctx.sync()
    })
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
  return await vueExcel.excel.run(async (ctx: Excel.RequestContext) => {
    let excelBinding

    if (props.address) {
      excelBinding = ctx.workbook.bindings.add(
        props.address, Excel.BindingType.range, props.address
      )
    } else if (props.name) {
      excelBinding = ctx.workbook.bindings.addFromNamedItem(
        props.name, Excel.BindingType.range, props.name
      )
    } else {
      console.error("You must pass either address or name props to a Range component!")
      return
    }

    excelBinding.onDataChanged.add(async (event: Excel.BindingDataChangedEventArgs) => {
      emit('dataChanged', event)
    })

    excelBinding.onSelectionChanged.add(async (event: Excel.BindingSelectionChangedEventArgs) => {
      emit('selectionChanged', event)
    })

    await ctx.sync()

    binding.value = excelBinding
  })
})

</script>

<template>
</template>