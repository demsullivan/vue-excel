import { type Component, type DefineComponent, type EmitsOptions } from 'vue'
import type Context from '@/Context'

export type RouteComponent = Component | DefineComponent

export type PluginOptions = {
  prefix?: string
}

export type NormalizedRoute = {
  activated(ctx: Context, worksheet: Excel.Worksheet): Promise<boolean>
  component: RouteComponent
  props: string[]
}

export type RouteWithSheetName = { sheetName: string; component: DefineComponent; props?: string[] }
export type RouteWithNamedRef = {
  namedRef: string
  value: string
  component: DefineComponent
  props?: string[]
}

export type Route =
  | RouteWithSheetName
  | RouteWithNamedRef
  | (Pick<NormalizedRoute, 'activated' | 'component'> & Partial<NormalizedRoute>)
