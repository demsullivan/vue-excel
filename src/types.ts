import { type Component, type DefineComponent, type EmitsOptions } from 'vue'

export type MaybeComponent = Component | DefineComponent

export type ComponentRegistry = Record<string, { component: MaybeComponent, props: Record<string, any>}>

export type ComponentList = Component[] | ComponentRegistry

export type PluginOptions = {
  components: ComponentList,
  workbookEmits?: EmitsOptions
}