import type { NormalizedRoute, Route, RouteWithSheetName, RouteWithNamedRef } from './types'
import type Context from './Context'

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

export default function normalizeRoutes(routes: Route[]): NormalizedRoute[] {
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
