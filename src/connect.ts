import { type App } from 'vue'
import { VueExcelGlobalState } from './state'
import { type Route, type PluginOptions } from '@/types'
import normalizeRoutes from '@/routes'
import Workbook from '@/components/Workbook.vue'
import Worksheet from '@/components/Worksheet.vue'
import Range from '@/components/Range.vue'
import Table from '@/components/Table.vue'

function installComponents(app: App, options: PluginOptions): void {
  const prefix = options.prefix || ''
  app
    .component(`${prefix}Workbook`, Workbook)
    .component(`${prefix}Worksheet`, Worksheet)
    .component(`${prefix}Range`, Range)
    .component(`${prefix}Table`, Table)
}

export async function connectExcel(routes: Route[] = []) {
  if (!window.Office) {
    throw 'Office could not be found! Are you sure you loaded office.js in your index.html?'
  }
  if (!Excel) {
    throw 'Excel could not be found!'
  }

  const normalizedRoutes = normalizeRoutes(routes)

  return await Excel.run(async (ctx: Excel.RequestContext) => {
    return {
      install(app: App, options: PluginOptions): void {
        installComponents(app, options)

        const state = new VueExcelGlobalState(ctx, normalizedRoutes)
        app.provide('vueExcel', state)
      }
    }
  })
}
