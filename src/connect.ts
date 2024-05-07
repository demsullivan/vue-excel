import { type App } from 'vue'
import { VueExcelGlobalState } from './state'
import { type Route, type PluginOptions } from '@vue-excel/types'
import normalizeRoutes from '@vue-excel/routes'
import Workbook from '@vue-excel/components/Workbook.vue'
import Worksheet from '@vue-excel/components/Worksheet.vue'
import Range from '@vue-excel/components/Range.vue'
import Table from '@vue-excel/components/Table.vue'
import Taskpane from '@vue-excel/components/Taskpane.vue'

function installComponents(app: App, options: PluginOptions): void {
  const prefix = options.prefix || ''
  app
    .component(`${prefix}Workbook`, Workbook)
    .component(`${prefix}Worksheet`, Worksheet)
    .component(`${prefix}Range`, Range)
    .component(`${prefix}Table`, Table)
    .component(`${prefix}Taskpane`, Taskpane)
}

export async function connectExcel(routes: Route[] = []) {
  if (!window.Office) {
    throw 'Office could not be found! Are you sure you loaded office.js in your index.html?'
  }
  if (!window.Excel) {
    throw 'Excel could not be found!'
  }

  const normalizedRoutes = normalizeRoutes(routes)

  return await Excel.run(async (ctx: Excel.RequestContext) => {
    return {
      install(app: App, options: PluginOptions = {}): void {
        installComponents(app, options)

        const state = new VueExcelGlobalState(ctx, normalizedRoutes)
        app.provide('vueExcel', state)
      }
    }
  })
}
