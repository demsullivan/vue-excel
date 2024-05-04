import { shallowRef, ref, type EmitsOptions } from 'vue'
import { type Component, type Ref, type DefineComponent } from 'vue'
import { type ComponentList, type ComponentRegistry, type MaybeComponent, type PluginOptions } from "./types"
import { Context } from '.'

export default class VueExcel extends EventTarget {
  components: Ref<ComponentRegistry>
  activeComponent: Ref<Component | null>
  activeComponentProps: Ref<Record<string, any>>
  workbookProps: Ref<Record<string, any>>
  excel: typeof Excel
  workbook: Ref<Excel.Workbook | null>
  activeWorksheet: Ref<Excel.Worksheet | null>
  worksheets: Ref<Excel.WorksheetCollection | null>
  workbookEmits: EmitsOptions
  context!: Context

  constructor(options: PluginOptions) {
    super()
    let namedComponents = {}

    if (options.components) {
      namedComponents = this.normalizeComponents(options.components)
    }
    
    this.excel = Excel

    this.activeComponent = shallowRef(null)
    this.activeComponentProps = ref({})

    this.workbook = shallowRef(null)
    this.workbookProps = ref({})
    this.activeWorksheet = shallowRef(null)
    this.worksheets = shallowRef(null)
    this.components = shallowRef({})
    this.workbookEmits = options.workbookEmits || []

    window.Office.onReady(async () => {
      OfficeExtension.config.extendedErrorLogging = true
      await Excel.run(this.initialize.bind(this, namedComponents))
    })
  }

  normalizeComponents(components: ComponentList): Record<string, MaybeComponent> {
    if (components instanceof Array) {
      return components.reduce((registry: Record<string, MaybeComponent>, component: Component) => {
        const defineComponent = component as DefineComponent
        if (defineComponent.__name !== undefined) {
          registry[defineComponent.__name] = component
        }
        return registry
      }, {})
    } else {
      return components
    }
  }
  
  async initialize(components: Record<string, MaybeComponent>, ctx: Excel.RequestContext) {
    this.context = new Context(ctx)
    const workbook = ctx.workbook.load()
    const activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet().load()
    const workbookNames = ctx.workbook.names.load(['name', 'value'])
    const worksheets = ctx.workbook.worksheets.load(['names', 'name'])

    ctx.workbook.worksheets.onActivated.add(this.worksheetActivated.bind(this))
    
    await ctx.sync()

    this.workbook.value = workbook
    this.activeWorksheet.value = activeWorksheet
    this.worksheets.value = worksheets

    this.components.value = worksheets.items.reduce((registry: ComponentRegistry, sheet: Excel.Worksheet) => {
      const componentName = sheet.names.items.find((item: Excel.NamedItem) => item.name == 'vue__component')?.value
      const props: Record<string, any> = this.extractProps(sheet.names)

      registry[sheet.name] = { component: components[componentName], props: props }
      return registry
    }, {})

    this.workbookProps.value = this.extractProps(workbookNames)
    this.worksheetActivated({ worksheetId: this.activeWorksheet.value.name, type: 'WorksheetActivated' })
  }

  change(component: DefineComponent, props: Record<string, any>) {
    console.debug(`[VueExcel] Changing active component to ${component.__name}`)
    this.activeComponent.value = component
    this.activeComponentProps.value = props
  }

  async worksheetActivated(event: Excel.WorksheetActivatedEventArgs) {
    await Excel.run(async (ctx: Excel.RequestContext) => {
      const worksheet = ctx.workbook.worksheets.getItem(event.worksheetId).load()
      await ctx.sync()
      this.activeWorksheet.value = worksheet

      let { component, props } = this.components.value[worksheet.name]
  
      if (component !== undefined) {
        this.change(component as DefineComponent, props)
      }
    })
  }

  extractProps(names: Excel.NamedItemCollection): Record<string, any> {
    return names.items.reduce((props: Record<string, any>, item: Excel.NamedItem) => {
      if (item.name.match(/^vue__props__/) !== null) {
        props[item.name.replace(/^vue__props__/, "")] = item.value
      }

      return props
    }, {})
  }
}