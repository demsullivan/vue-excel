import { type Component, type DefineComponent } from 'vue'
import type Context from './Context'

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

export enum ExcelCellValueType {
  array = 'Array',
  boolean = 'Boolean',
  double = 'Double',
  entity = 'Entity',
  empty = 'Empty',
  error = 'Error',
  formattedNumber = 'FormattedNumber',
  linkedEntity = 'LinkedEntity',
  reference = 'Reference',
  string = 'String',
  notAvailable = 'NotAvailable',
  webImage = 'WebImage'
}
