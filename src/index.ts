import { shallowRef, type App, type DefineComponent, type ShallowRef } from 'vue'
import {
  type PluginOptions,
  type Route,
  type NormalizedRoute,
  type RouteWithSheetName,
  type RouteWithNamedRef
} from './types'
import Context from './Context'
import Workbook from './components/Workbook.vue'
import Worksheet from './components/Worksheet.vue'
import Range from './components/Range.vue'
import Table, { type TableChangedEvent } from './components/Table.vue'

export { Context, Workbook, Worksheet, Range, Table }
export type { TableChangedEvent }

function routeActivationHandlerBySheetName(
  sheetName: string
): (ctx: Context, worksheet: Excel.Worksheet) => Promise<boolean> {
  return async (ctx: Context, worksheet: Excel.Worksheet) => {
    return worksheet.name == sheetName
  }
}

function routeActivationHandlerByNamedRef(
  namedRef: string,
  value: string
): (ctx: Context, worksheet: Excel.Worksheet) => Promise<boolean> {
  return async (ctx: Context, worksheet: Excel.Worksheet) => {
    const names = worksheet.names.load(['name', 'value'])
    await ctx.sync()
    const namedValue = names.items.find((name) => name.name == namedRef)

    if (namedValue) {
      return namedValue.value == value
    } else {
      return false
    }
  }
}

function normalizeRoutes(routes: Route[]): NormalizedRoute[] {
  return routes.map((route: Route) => {
    if ((route as RouteWithSheetName).sheetName) {
      const sheetNameRoute = <RouteWithSheetName>route
      return <NormalizedRoute>{
        component: route.component,
        activated: routeActivationHandlerBySheetName(sheetNameRoute.sheetName),
        props: route.props || []
      }
    } else if ((route as RouteWithNamedRef).namedRef) {
      const namedRefRoute = <RouteWithNamedRef>route
      return <NormalizedRoute>{
        component: route.component,
        activated: routeActivationHandlerByNamedRef(namedRefRoute.namedRef, namedRefRoute.value),
        props: route.props || []
      }
    } else {
      return <NormalizedRoute>route
    }
  })
}

function installComponents(app: App, options: PluginOptions): void {
  const prefix = options.prefix || ''
  app
    .component(`${prefix}Workbook`, Workbook)
    .component(`${prefix}Worksheet`, Worksheet)
    .component(`${prefix}Range`, Range)
    .component(`${prefix}Table`, Table)
}

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
