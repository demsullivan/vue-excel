import { type App, type EmitsOptions, type ComponentOptions, defineComponent } from 'vue'
import { type ComponentList, type PluginOptions, type MaybeComponent, type ComponentRegistry } from './types'
import VueExcel from './VueExcel'
import Context from './Context'
import createWorkbookComponent from './components/Workbook.vue'
import Worksheet from './components/Worksheet.vue'
import Range from './components/Range.vue'

export { VueExcel, Context }

export function connectExcel() {
  return {
    install(app: App, options: PluginOptions): void {
      if (!window.Office) { throw "No office! Bad!" }
      const excel = new VueExcel(options)

      this.installComponents(app, options)
      this.installExcel(app, excel)
    },

    installComponents(app: App, options: PluginOptions): void {
      app.component('Workbook', createWorkbookComponent(options.workbookEmits))
        .component('Worksheet', Worksheet)
        .component('Range', Range)
    },

    installExcel(app: App, excel: VueExcel): void {
      app.provide('vueExcel', excel)
      app.config.globalProperties.vueExcel = excel
    }
  }
}