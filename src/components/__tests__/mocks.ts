import OfficeAddinMock from 'office-addin-mock'

type EventHandlerMock = {
  listener: (event: any) => void
  add(callback: (event: any) => void): void
  remove(callback: (event: any) => void): void
  fire(event: any): void
}

export function createEventHandlerMock(): EventHandlerMock {
  return {
    listener: () => {},
    add(callback: (event: any) => void) {
      this.listener = callback
    },
    remove(callback: (event: any) => void) {
      this.listener = () => {}
    },
    fire(event: any) {
      this.listener(event)
    }
  }
}

export function createContextMock(properties: Record<string, any> = {}) {
  const mockData = {
    workbook: {
      name: 'Test Workbook.xlsx',
      worksheets: {
        activeWorksheet: createWorksheetMock({ name: 'Sheet1' }),
        getActiveWorksheet() {
          return this.activeWorksheet
        },
        onActivated: createEventHandlerMock(),
        ...(properties.worksheets || {})
      },
      ...(properties.workbook || {})
    }
  }

  mockData.workbook.worksheets.items = [mockData.workbook.worksheets.activeWorksheet]

  return new OfficeAddinMock.OfficeMockObject(mockData) as any
}

export function createWorksheetMock(properties: {}) {
  return {
    id: '1',
    name: 'Sheet1',
    names: {},
    onCalculated: createEventHandlerMock(),
    onChanged: createEventHandlerMock(),
    onColumnSorted: createEventHandlerMock(),
    onFormatChanged: createEventHandlerMock(),
    onFormulaChanged: createEventHandlerMock(),
    onRowHiddenChanged: createEventHandlerMock(),
    onRowSorted: createEventHandlerMock(),
    onSelectionChanged: createEventHandlerMock(),
    onSingleClicked: createEventHandlerMock(),
    onVisibilityChanged: createEventHandlerMock(),
    ...properties
  }
}
