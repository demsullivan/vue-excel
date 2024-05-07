<script setup lang="ts">
import { computed, inject, onMounted, ref, shallowRef, watch } from 'vue'
import type { NormalizedRoute, RouteComponent } from '@/types'
import type { VueExcelGlobalState } from '..'

type RoutedComponentRegistry = Record<string, { component: RouteComponent; props: Record<string, any> }>

const vueExcel = inject('vueExcel') as VueExcelGlobalState

const context = vueExcel.context
const routes = vueExcel.routes

const workbookNames = shallowRef<Excel.NamedItemCollection>()
const activeWorksheet = shallowRef<Excel.Worksheet>()
const activeWorksheetNames = shallowRef<Excel.NamedItemCollection>()
const computedRoutes = ref<any[]>([])

async function worksheetActivated(event: Excel.WorksheetActivatedEventArgs) {
  const { xlSheet, xlNames } = await context.fetch(async (ctx: Excel.RequestContext) => {
    const xlSheet = ctx.workbook.worksheets.getItem(event.worksheetId)
    return {
      xlSheet,
      xlNames: xlSheet.names
    }
  })
  vueExcel.activeWorksheet.value = xlSheet
  activeWorksheetNames.value = xlNames
}

onMounted(async () => {
  const { xlWorkbook, xlWorksheets, xlNames, xlActiveWorksheet } = await context.fetch(
    async (ctx: Excel.RequestContext) => {
      const xlWorksheets = ctx.workbook.worksheets

      xlWorksheets.onActivated.add(worksheetActivated)

      return {
        xlWorksheets,
        xlWorkbook: ctx.workbook,
        xlNames: ctx.workbook.names,
        xlActiveWorksheet: ctx.workbook.worksheets.getActiveWorksheet()
      }
    }
  )

  vueExcel.workbook.value = xlWorkbook
  vueExcel.worksheets.value = xlWorksheets
  vueExcel.activeWorksheet.value = xlActiveWorksheet
  workbookNames.value = xlNames
})

const routedComponents = computed<RoutedComponentRegistry>(() => {
  if (!vueExcel.worksheets.value) return {}

  return vueExcel.worksheets.value.items.reduce((registry: RoutedComponentRegistry, worksheet: Excel.Worksheet) => {
    const route = routes.find((route) => route.activated(context, worksheet))
    if (route) {
      registry[worksheet.name] = {
        component: route.component,
        props: { worksheet }
      }
    }
    return registry
  }, {})
})
</script>

<template>
  <div v-if="vueExcel.workbook">
    <slot></slot>
    <div
      v-for="(routedComponent, worksheet) in routedComponents"
      :style="worksheet == vueExcel.activeWorksheet.value?.name ? null : 'display: none'"
    >
      <component :is="routedComponent.component" v-bind="routedComponent.props" />
    </div>
  </div>
</template>
