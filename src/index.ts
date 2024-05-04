import { type App, type EmitsOptions, type ComponentOptions, defineComponent } from 'vue'
import { type ComponentList, type PluginOptions, type MaybeComponent, type ComponentRegistry } from './types'
import VueExcel from './VueExcel'
import Context from './Context'
import createWorkbookComponent from './components/Workbook.vue'
import Worksheet from './components/Worksheet.vue'
import Range from './components/Range.vue'
import Table, { type TableChangedEvent } from './components/Table.vue'

export { VueExcel, Context, Worksheet, Range, Table }
export type { TableChangedEvent }

function installComponents(app: App, options: PluginOptions): void {
  const prefix = options.prefix || ""
  app.component(`${prefix}Workbook`, createWorkbookComponent(options.workbookEmits))
    .component(`${prefix}Worksheet`, Worksheet)
    .component(`${prefix}Range`, Range)
    .component(`${prefix}Table`, Table)
}

function installExcel(app: App, excel: VueExcel): void {
  app.provide('vueExcel', excel)
  app.config.globalProperties.vueExcel = excel
}

export async function connectExcel() {
  if (!window.Office) { throw "Office could not be found! Are you sure you loaded office.js in your index.html?" }
  if (!Excel) { throw "Excel could not be found!" }

  return await Excel.run(async (ctx: Excel.RequestContext) => {
    return {
      install(app: App, options: PluginOptions): void {
        const excel = new VueExcel(options, ctx)
  
        installComponents(app, options)
        installExcel(app, excel)
      }
    }
  })
}