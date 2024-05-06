import Worksheet from '@/components/Worksheet.vue'
import { mount, flushPromises } from '@vue/test-utils'
import { expect, beforeEach } from 'vitest'
import { type GlobalStateFixture, composeTestWithState } from './utils'
import { createContextMock, createWorksheetMock } from './mocks'
import { defineComponent, h, inject, nextTick } from 'vue'

const mockContext = createContextMock({
  worksheets: {
    Sheet2: createWorksheetMock({ name: 'Sheet2' }),
    getItem(name: 'Sheet1' | 'Sheet2') {
      if (name == 'Sheet1') return this.activeWorksheet
      return this[name]
    }
  }
})

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

  mockContext.workbook.worksheets.activeWorksheet.onChanged.fire({ address: 'A1' })
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
