import { VueExcelGlobalState } from 'vue-excel/state'
import { test } from 'vitest'

export interface GlobalStateFixture {
  globalState: VueExcelGlobalState
}

export function composeTestWithState(contextMock: any) {
  return test.extend<GlobalStateFixture>({
    globalState: async ({}, use) => {
      const globalState = new VueExcelGlobalState(contextMock)
      await use(globalState)
    }
  })
}
