<script setup lang="ts">
import { inject, onMounted, ref, shallowRef, watch } from 'vue'
import type { NormalizedRoute } from '@/types'
import type { VueExcelGlobalState } from '..'

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

watch(
  () => vueExcel.activeWorksheet.value,
  async (newValue) => {
    if (!newValue) return

    computedRoutes.value = await Promise.all(
      routes.map(async (route: NormalizedRoute) => {
        return {
          isActive: newValue ? await route.activated(context, newValue) : false,
          component: route.component
        }
      })
    )
  }
)
</script>

<template>
  <div v-if="vueExcel.workbook">
    <slot></slot>
  </div>
  <div v-for="route in computedRoutes" :style="route.isActive ? null : 'display: none'">
    <component :is="route.component" />
  </div>
</template>
