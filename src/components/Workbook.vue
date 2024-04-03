<script lang="ts">
import { type EmitsOptions, defineComponent, type ComponentOptions, h } from 'vue';
import { type MaybeComponent, type ComponentRegistry } from '../types'
import _ from 'lodash'

export default function(emits?: EmitsOptions) {
  return defineComponent({
    inject: ['vueExcel'],
    emits: emits,
    methods: {
      eventsForComponent(component: MaybeComponent) {
        const componentWithEmits = component as ComponentOptions
        if (componentWithEmits == null || componentWithEmits.emits == undefined) return {}
      
        return componentWithEmits.emits.reduce((events: Record<string, Function>, emitName: string) => {
          events[`on${_.upperFirst(emitName)}`] = (event: any) => {
            this.$emit(emitName, event)
          }
      
          return events
        }, {})
      }
    },
    computed: {
      activeWorksheetName(): string { return this.vueExcel.activeWorksheet.value.name },
      components(): ComponentRegistry { return this.vueExcel.components.value }
    },
    render() {
      const componentNodes = Object.keys(this.components).map(name => {
        const config = this.components[name]

        if (config.component === undefined) return
        return h(
          'div',
          { style: name !== this.activeWorksheetName ? 'display: none' : '' },
          [
            h(
              config.component,
              { name: name, ...config.props, ...this.eventsForComponent(config.component) }
              )
          ]
        )
      });

      return [
        ...componentNodes,
        this.$slots.default()
      ]
    }
  });
}
</script>