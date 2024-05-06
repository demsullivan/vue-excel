function addEventHandler(eventMocks: EventTarget, name: string) {
  return {
    add(callback: () => void) {
      eventMocks.addEventListener(name, callback)
    },

    remove(callback: () => void) {
      eventMocks.removeEventListener(name, callback)
    }
  }
}

export function worksheet(properties: {}, eventTarget: EventTarget = new EventTarget()) {
  return {
    id: '1',
    name: 'Sheet1',
    names: {},
    onCalculated: addEventHandler(eventTarget, 'onCalculated'),
    onChanged: addEventHandler(eventTarget, 'onChanged'),
    onColumnSorted: addEventHandler(eventTarget, 'onColumnSorted'),
    onFormatChanged: addEventHandler(eventTarget, 'onFormatChanged'),
    onFormulaChanged: addEventHandler(eventTarget, 'onFormulaChanged'),
    onRowHiddenChanged: addEventHandler(eventTarget, 'onRowHiddenChanged'),
    onRowSorted: addEventHandler(eventTarget, 'onRowSorted'),
    onSelectionChanged: addEventHandler(eventTarget, 'onSelectionChanged'),
    onSingleClicked: addEventHandler(eventTarget, 'onSingleClicked'),
    onVisibilityChanged: addEventHandler(eventTarget, 'onVisibilityChanged'),
    ...properties
  }
}
