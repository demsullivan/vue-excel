import Worksheet from '@/components/Worksheet.vue'
import OfficeAddinMock from 'office-addin-mock'
import { mount, flushPromises } from '@vue/test-utils'
import { expect, beforeEach } from 'vitest'
import { type GlobalStateFixture, composeTestWithState } from './utils'
import { worksheet } from './mocks'
import { defineComponent, h, inject, nextTick } from 'vue'

const activeWorksheetEventTarget = new EventTarget()

const mockData = {
  workbook: {
    worksheets: {
      Sheet2: worksheet({ name: 'Sheet2' }),
      activeWorksheet: worksheet({ name: 'Sheet1' }, activeWorksheetEventTarget),
      getItem(name: 'Sheet1' | 'Sheet2') {
        if (name == 'Sheet1') return this.activeWorksheet
        return this[name]
      }
    }
  }
}

const mockContext = new OfficeAddinMock.OfficeMockObject(mockData) as any
const it = composeTestWithState(mockContext)

beforeEach<GlobalStateFixture>(async ({ globalState }) => {
  const sheet = mockContext.workbook.worksheets.activeWorksheet.load()
  await mockContext.sync()
  globalState.activeWorksheet.value = sheet
})

it('does not render slots if the worksheet is not active', async ({ globalState }) => {
  const wrapper = mount(Worksheet, {
    props: {
      name: 'Sheet2'
    },
    global: { provide: { vueExcel: globalState } }
  })

  expect(wrapper.html()).toBe('<!--v-if-->')
})

it('renders slots if the worksheet is active', async ({ globalState }) => {
  const wrapper = mount(Worksheet, {
    props: {
      name: 'Sheet1'
    },
    slots: { default: 'Hello' },
    global: { provide: { vueExcel: globalState } }
  })

  await nextTick()
  expect(wrapper.html()).toContain('Hello')
})

it('emits the onChanged event', async ({ globalState }) => {
  const wrapper = mount(Worksheet, {
    props: {
      name: 'Sheet1'
    },
    slots: { default: 'Hello' },
    global: { provide: { vueExcel: globalState } }
  })

  await flushPromises()

  const event = new CustomEvent('onChanged', { detail: { address: 'A1' } })
  activeWorksheetEventTarget.dispatchEvent(event)
  expect(wrapper.emitted()).toHaveProperty('changed')
})

it('provides vueExcel.scope.worksheet', async ({ globalState }) => {
  const StubComponent = defineComponent({
    template: '<div id="worksheet-name">{{ worksheet?.name }}</div>',
    setup(props) {
      const worksheet = inject('vueExcel.scope.worksheet')
      return { worksheet }
    }
  })

  const wrapper = mount(Worksheet, {
    props: { name: 'Sheet1' },
    slots: { default: () => h(StubComponent) },
    global: { provide: { vueExcel: globalState } }
  })

  await flushPromises()

  expect(wrapper.find('#worksheet-name').text()).toBe('Sheet1')
})
