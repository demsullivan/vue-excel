import { expect, it, vi } from 'vitest'
import Context from '@/Context'

import OfficeAddinMock from 'office-addin-mock'
import test, { describe } from 'node:test'

const mockData = {
  workbook: {
    name: 'Test Workbook.xlsx',
    worksheets: {
      activeWorksheet: {
        name: 'Sheet1'
      },
      getActiveWorksheet() {
        return this.activeWorksheet
      }
    },
    range: {
      address: 'C2:G3',
      values: [
        ['Hello', 'World', 'From', 'Vue', 'Excel'],
        ['Hello', 'World', 'From', 'Vue', 'Excel']
      ]
    },
    getSelectedRange() {
      return this.range
    }
  }
}

const contextMock = new OfficeAddinMock.OfficeMockObject(mockData)

const subject = new Context(contextMock)

describe('.fetch', () => {
  it('automatically loads the objects', async () => {
    const { range } = await subject.fetch(async (ctx) => {
      return {
        range: ctx.workbook.getSelectedRange()
      }
    })

    expect(range.address).toBe('C2:G3')
  })

  it('throws errors when unloaded properties are accessed', async () => {
    // Not exactly sure why OfficeMockObject doesn't throw an error here...
    expect(contextMock.workbook.name).toEqual('Error, property was not loaded')
  })
})

describe('.sync', () => {
  it('calls sync on the RequestContext', () => {
    const spy = vi.spyOn(contextMock, 'sync')

    subject.sync()

    expect(spy).toHaveBeenCalledTimes(1)
  })
})

describe('.run', () => {
  it('automatically syncs the context', async () => {
    const { range } = await subject.fetch(async (ctx) => ({
      range: ctx.workbook.getSelectedRange()
    }))

    const spy = vi.spyOn(contextMock, 'sync')

    await subject.run(async (ctx) => {
      range.values = [
        ['C', 'D', 'E', 'F', 'G'],
        ['C', 'D', 'E', 'F', 'G']
      ]
    })

    expect(spy).toHaveBeenCalledTimes(1)
  })
})
