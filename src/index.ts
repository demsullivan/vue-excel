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

export function connectExcel() {
  return {
    install(app: App, options: PluginOptions): void {
      if (!window.Office) { throw "No office! Bad!" }
      const excel = new VueExcel(options)

      this.installComponents(app, options)
      this.installExcel(app, excel)
    },

    installComponents(app: App, options: PluginOptions): void {
      const prefix = options.prefix || ""
      app.component(`${prefix}Workbook`, createWorkbookComponent(options.workbookEmits))
        .component(`${prefix}Worksheet`, Worksheet)
        .component(`${prefix}Range`, Range)
        .component(`${prefix}Table`, Table)
    },

    installExcel(app: App, excel: VueExcel): void {
      app.provide('vueExcel', excel)
      app.config.globalProperties.vueExcel = excel
    }
  }
}