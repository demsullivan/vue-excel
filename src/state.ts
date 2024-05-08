import { type ShallowRef, shallowRef } from 'vue'
import Context from './Context'
import { type NormalizedRoute } from './types'

export class VueExcelGlobalState {
  context: Context
  routes: NormalizedRoute[]
  workbook: ShallowRef<Excel.Workbook | undefined>
  worksheets: ShallowRef<Excel.WorksheetCollection | undefined>
  activeWorksheet: ShallowRef<Excel.Worksheet | undefined>

  constructor(context: Excel.RequestContext, routes: NormalizedRoute[] = []) {
    this.context = new Context(context)
    this.routes = routes
    this.workbook = shallowRef()
    this.worksheets = shallowRef()
    this.activeWorksheet = shallowRef()
  }
}
