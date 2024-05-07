import { expect, it, beforeEach, vi, describe } from 'vitest'
import { flushPromises, mount } from '@vue/test-utils'
import { connectExcel } from '@vue-excel/index'
import type { RouteComponent } from '@vue-excel/types'
import { defineComponent, inject, nextTick } from 'vue'
import { createContextMock } from '@vue-excel/components/__tests__/mocks'

let vueExcel: any

const mockContext = createContextMock()

async function componentWithPlugin(Component: RouteComponent) {
  const wrapper = mount(Component, {
    global: {
      plugins: [vueExcel]
    }
  })

  await flushPromises()

  return wrapper
}

describe('when Office is not present', () => {
  it('throws an error', async () => {
    await expect(connectExcel()).rejects.toEqual(
      'Office could not be found! Are you sure you loaded office.js in your index.html?'
    )
  })
})

describe('when Excel is not present', () => {
  beforeEach(() => {
    vi.stubGlobal('Office', {})
  })

  it('throws an error', async () => {
    await expect(connectExcel()).rejects.toEqual('Excel could not be found!')
  })
})

describe('vueExcel', () => {
  beforeEach(async () => {
    vi.stubGlobal('Office', {})
    vi.stubGlobal('Excel', {
      run: async (callback: any) => {
        return await callback(mockContext)
      }
    })

    vueExcel = await connectExcel()
  })

  it('provides vueExcel', async () => {
    const StubComponent = defineComponent({
      template: '<Workbook><div id="workbook-name">{{ vueExcel.workbook.name }}</div></Workbook>',
      setup(props) {
        const vueExcel = inject('vueExcel')
        return { vueExcel }
      }
    })

    const wrapper = await componentWithPlugin(StubComponent)

    expect(wrapper.vm.vueExcel.workbook.value.name).toEqual('Test Workbook.xlsx')
  })
})
