import Workbook from '../Workbook.vue'
import Context from 'vue-excel/Context'
import { VueExcelGlobalState } from 'vue-excel/state'
import { flushPromises, mount } from '@vue/test-utils'
import { expect, it, beforeEach, describe } from 'vitest'
import { h } from 'vue'
import { createContextMock, createWorksheetMock, createEventHandlerMock } from './mocks'

const mockContext = createContextMock({
  worksheets: {
    sheetOne: createWorksheetMock({ id: '1', name: 'Sheet1' }),
    sheetTwo: createWorksheetMock({ id: '2', name: 'Sheet2' }),
    activeWorksheet: createWorksheetMock({ id: '1', name: 'Sheet1' }),
    getItem(id: string) {
      if (id == '1') return this.sheetOne
      if (id == '2') return this.sheetTwo
    },
    onActivated: createEventHandlerMock(),
    getActiveWorksheet() {
      return this.activeWorksheet
    }
  }
})

interface LocalTestContext {
  globalState: VueExcelGlobalState
}

beforeEach<LocalTestContext>(async (context) => {
  context.globalState = new VueExcelGlobalState(mockContext)
})

async function mountWorkbook({
  globalState,
  flush = true,
  slots = {}
}: { globalState?: VueExcelGlobalState; flush?: boolean; slots?: Record<string, any> } = {}) {
  const wrapper = mount(Workbook, {
    slots,
    global: {
      provide: {
        vueExcel: globalState
      }
    }
  })

  if (flush) await flushPromises()

  return wrapper
}

it<LocalTestContext>('sets workbook on the global state object', async ({ globalState }) => {
  await mountWorkbook({ globalState })
  expect(globalState.workbook.value?.name).toEqual('Test Workbook.xlsx')
})

it<LocalTestContext>('sets worksheets on the global state object', async ({ globalState }) => {
  await mountWorkbook({ globalState })
  expect(globalState.worksheets.value).toEqual(mockContext.workbook.worksheets)
})

it<LocalTestContext>('sets the active worksheet on the global state object', async ({ globalState }) => {
  await mountWorkbook({ globalState })
  expect(globalState.activeWorksheet.value).toEqual(mockContext.workbook.worksheets.activeWorksheet)
})

it<LocalTestContext>('updates activeWorksheet when the event listener is called', async ({ globalState }) => {
  await mountWorkbook({ globalState })

  mockContext.workbook.worksheets.onActivated.fire({ worksheetId: '2' })
  await flushPromises()
  expect(globalState.activeWorksheet.value?.name).toEqual('Sheet2')
})

describe('Advanced Routing', () => {
  beforeEach<LocalTestContext>(async (context) => {
    const routes = [
      {
        async activated(ctx: Context, worksheet: Excel.Worksheet) {
          return worksheet.name == 'Sheet1'
        },
        component: () => h('div', 'Component 1'),
        props: []
      },
      {
        async activated(ctx: Context, worksheet: Excel.Worksheet) {
          return worksheet.name == 'Sheet2'
        },
        component: () => h('div', 'Component 2'),
        props: []
      }
    ]

    context.globalState = new VueExcelGlobalState(mockContext, routes)
  })

  it<LocalTestContext>('activates the first route', async ({ globalState }) => {
    const wrapper = await mountWorkbook({ globalState })

    expect(wrapper.html()).toContain('<div>\n  <div>Component 1')
  })

  it<LocalTestContext>('hides the second route', async ({ globalState }) => {
    const wrapper = await mountWorkbook({ globalState })

    expect(wrapper.html()).toContain('<div style="display: none;">\n  <div>Component 2')
  })

  it<LocalTestContext>('activates the second route on worksheet change', async ({ globalState }) => {
    const wrapper = await mountWorkbook({ globalState })

    mockContext.workbook.worksheets.onActivated.fire({ worksheetId: '2' })
    await flushPromises()

    expect(wrapper.html()).toContain('<div>\n  <div>Component 2')
  })

  it<LocalTestContext>('hides the first route on worksheet change', async ({ globalState }) => {
    const wrapper = await mountWorkbook({ globalState })

    mockContext.workbook.worksheets.onActivated.fire({ worksheetId: '2' })
    await flushPromises()

    expect(wrapper.html()).toContain('<div style="display: none;">\n  <div>Component 1')
  })
})
