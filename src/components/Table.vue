<script setup lang="ts">
import { type ShallowRef, inject, shallowRef, watch, ref } from 'vue'
import type Context from '../Context'
import type { VueExcelGlobalState } from '../state'

export type TableRowRecord = Record<string, any>

export interface TableRow {
  rowIndex: number
  range: Excel.Range
  record: TableRowRecord
}

export interface TableChangedEvent extends Excel.TableChangedEventArgs {
  range: Excel.Range
  tableRow: TableRow
}

type Props = {
  name?: string
  headers?: string[]
  data?: TableRowRecord[]
}

type Emits = {
  dataChanged: [event: TableChangedEvent]
  selectionChanged: [event: Excel.TableSelectionChangedEventArgs]
}

const props = defineProps<Props>()
const emit = defineEmits<Emits>()

const vueExcel = inject('vueExcel') as VueExcelGlobalState
const context = vueExcel.context

const worksheet = inject<ShallowRef<Excel.Worksheet>>('vueExcel.scope.worksheet')
const table = shallowRef<Excel.Table>()
const headerRange = shallowRef<Excel.Range>()
const headers = ref<string[] | undefined>(props.headers)

async function updateTable(overwrite: boolean) {
  if (!worksheet) return

  const { xlTable, xlHeaderRange } = await context.fetch(async (ctx: Excel.RequestContext) => {
    let excelTable

    if (table.value) {
      excelTable = table.value
    } else {
      excelTable = props.name ? worksheet.value.tables.getItem(props.name) : worksheet?.value.tables.getItemAt(0)
      excelTable.load('name')
      await ctx.sync()
    }

    if (!excelTable) return { xlTable: undefined, xlHeaderRange: undefined }

    const tableName = excelTable.name

    if (overwrite && props.data && props.headers) {
      excelTable.delete()
      let address = headerRange.value ? headerRange.value.address : 'A1'

      const newTable = worksheet.value.tables.add(address, true)
      newTable.name = tableName
      newTable.getHeaderRowRange().values = [props.headers]

      props.data.forEach((record: Record<string, any>) => {
        const values = (props.headers as string[]).reduce(
          (data, key) => {
            data.push(record[key])
            return data
          },
          <string[]>[]
        )

        newTable.rows.add(undefined, [values])
      })

      excelTable = newTable
    }

    excelTable.onChanged.add(onDataChanged)
    excelTable.onSelectionChanged.add(onSelectionChanged)

    return {
      xlTable: excelTable,
      xlHeaderRange: excelTable.getHeaderRowRange()
    }
  })

  table.value = xlTable
  headerRange.value = xlHeaderRange

  if (!headers.value && xlHeaderRange) {
    headers.value = xlHeaderRange.values[0]
  }
}

function initializeTable(table: Excel.Table) {
  table.onChanged.add(onDataChanged)
  table.onSelectionChanged.add(onSelectionChanged)
}

watch(
  () => worksheet?.value,
  async (value) => {
    if (value) {
      updateTable(false)
    }
  }
)

watch(
  () => props.data,
  async (value) => {
    if (value == null || value == undefined) return
    if (!table.value) return
    if (!worksheet?.value) return
    if (!headerRange?.value) return

    updateTable(true)
  },
  { immediate: true }
)

////////////////////////////////////////////////
// PUBLIC API
////////////////////////////////////////////////
function getRowById(id: string | number) {
  // TODO
}

type GetRecordArgs = {
  id?: any
  row?: number
  field?: string
  value?: any
  range?: Excel.Range
}

async function getTableRow({ id }: { id: any }): Promise<TableRow | null>
async function getTableRow({ row }: { row: number }): Promise<TableRow | null>
async function getTableRow({ range }: { range: Excel.Range }): Promise<TableRow | null>
async function getTableRow({ field, value }: { field: string; value: any }): Promise<TableRow | null>
async function getTableRow({ id, row, field, value, range }: GetRecordArgs): Promise<TableRow | null> {
  let rowValues: any[] = []

  if (!headers?.value) return null

  if (id) {
    field = 'id'
    value = id
  }

  try {
    if (field && value) {
      range = await getRowRange({ field, value })
    } else if (row) {
      range = await getRowRange({ row })
    }
  } catch (e) {
    throw `getRecord: Could not find range for row based on provided values`
  }

  if (!range) return null

  try {
    rowValues = range.values[0]
  } catch {
    range.load('values')
    await context.sync()
    rowValues = range.values[0]
  }

  // TODO: check types and convert as necessary, especially dates
  const record = <TableRowRecord>headers.value?.reduce(
    (rowData, key, index) => {
      rowData[key] = rowValues[index] == '' ? null : rowValues[index]
      return rowData
    },
    <TableRowRecord>{}
  )

  return {
    rowIndex: row || range.rowIndex,
    range,
    record
  }
}

type UpdateRowArgs = {
  id?: string | number
  row?: number
  field?: string
  value?: any
  record?: Record<string, any>
  tableRow?: TableRow
}

async function updateTableRow({ id, record }: { id: string | number; record: Record<string, any> }): Promise<void>
async function updateTableRow({ row, record }: { row: number; record: Record<string, any> }): Promise<void>
async function updateTableRow({
  field,
  value,
  record
}: {
  field: string
  value: any
  record: Record<string, any>
}): Promise<void>
async function updateTableRow({ tableRow }: { tableRow: TableRow }): Promise<void>
async function updateTableRow({ id, row, field, value, record, tableRow }: UpdateRowArgs) {
  if (!headers?.value) return
  if (!table.value) return

  if (!record && !tableRow) throw `updateTableRow: you must specify record or tableRow!`

  let rowRange: Excel.Range | undefined

  if (id) {
    field = 'id'
    value = id
  }

  try {
    if (field && value) {
      rowRange = await getRowRange({ field, value })
    } else if (row) {
      rowRange = await getRowRange({ row })
    } else if (tableRow) {
      rowRange = tableRow.range
    }
  } catch (e) {
    throw `updateRow: Could not find range for row based on provided values`
  }

  if (!rowRange) return

  if (tableRow) record = tableRow.record

  record = <TableRowRecord>record

  rowRange.values = [
    headers.value.map((headerName: string) => {
      return Object.keys(record).includes(headerName) ? record[headerName] : ''
    })
  ]

  await context.sync()
}

defineExpose({
  getTableRow,
  updateTableRow
})

////////////////////////////////////////////////
// INTERNAL FUNCTIONS
////////////////////////////////////////////////
async function getRowRange({ field, value }: { field: string; value: any }): Promise<Excel.Range>
async function getRowRange({ row }: { row: number }): Promise<Excel.Range>
async function getRowRange({
  row,
  field,
  value
}: {
  row?: number
  field?: string
  value?: any
}): Promise<Excel.Range | undefined> {
  if (!worksheet?.value) return
  if (!headers?.value) return
  if (!headerRange?.value) return

  let rowIndex: number | null = null
  let rowRange: Excel.Range | undefined

  if (field && value) {
    rowIndex = await getRowByValue(field, value)
  } else if (row) {
    rowIndex = row
  }

  if (!rowIndex) throw `Could not find range`

  const { range } = await context.fetch(async () => {
    if (!headerRange.value) return { range: undefined }
    return {
      range: worksheet.value.getRangeByIndexes(rowIndex, 0, 1, headerRange.value.columnCount)
    }
  })

  return range
}

async function getRowByValue(field: string, value: any): Promise<number | null> {
  if (!table.value) return null

  const { foundCell } = await context.fetch(async (ctx: Excel.RequestContext) => {
    const columnRange = table.value?.columns.getItem(field).getRange()
    return {
      foundCell: columnRange?.findOrNullObject(value, {
        completeMatch: true,
        searchDirection: Excel.SearchDirection.forward
      })
    }
  })

  if (!foundCell) return null

  return foundCell.isNullObject ? null : foundCell?.rowIndex
}

////////////////////////////////////////////////
// EVENT HANDLERS
////////////////////////////////////////////////
async function onDataChanged(event: Excel.TableChangedEventArgs) {
  if (!table.value) return
  if (!worksheet?.value) return

  const range = event.getRangeOrNullObject(worksheet.value.context).load()

  await worksheet.value.context.sync()

  if (range.isNullObject) return

  const tableRow = await getTableRow({ row: range.rowIndex })

  if (!tableRow) return

  emit('dataChanged', { ...event, range, tableRow })
}

async function onSelectionChanged(event: Excel.TableSelectionChangedEventArgs) {
  emit('selectionChanged', event)
}
</script>

<template></template>
