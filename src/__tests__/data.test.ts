import { expect, it, describe, vi, beforeEach } from 'vitest'
import moment from 'moment-msdate'
import { ExcelCellValueType } from 'vue-excel/types'
import { cellToExcel, cellFromExcel, FormattedNumber } from 'vue-excel/data'

describe('from excel', () => {
  it('converts a string', () => {
    const cellValue = cellFromExcel({ type: ExcelCellValueType.string, basicValue: 'hello' })
    expect(cellValue).toEqual('hello')
  })

  it('converts a number', () => {
    const cellValue = cellFromExcel({ type: ExcelCellValueType.double, basicValue: 123 })
    expect(cellValue).toEqual(123)
  })

  it('converts a boolean', () => {
    const cellValue = cellFromExcel({ type: ExcelCellValueType.boolean, basicValue: true })
    expect(cellValue).toEqual(true)
  })

  it('converts an empty cell', () => {
    const cellValue = cellFromExcel({ type: ExcelCellValueType.empty })
    expect(cellValue).toEqual(null)
  })

  it('converts a formatted number', () => {
    const cellValue = cellFromExcel({
      type: ExcelCellValueType.formattedNumber,
      basicValue: 123,
      numberFormat: '0.00'
    })
    expect(cellValue).toEqual(new FormattedNumber(123, '0.00'))
  })

  it('returns the raw Excel CellValue if the type is unknown', () => {
    const cellValue = cellFromExcel({ type: ExcelCellValueType.entity, text: 'hello!' })
    expect(cellValue).toEqual({ type: ExcelCellValueType.entity, text: 'hello!' })
  })
})

describe('to excel', () => {
  it('converts a string', () => {
    const cellValue = cellToExcel('hello')
    expect(cellValue).toEqual({ type: ExcelCellValueType.string, basicValue: 'hello' })
  })

  it('converts a number', () => {
    const cellValue = cellToExcel(123)
    expect(cellValue).toEqual({ type: ExcelCellValueType.double, basicValue: 123 })
  })

  it('converts a boolean', () => {
    const cellValue = cellToExcel(true)
    expect(cellValue).toEqual({ type: ExcelCellValueType.boolean, basicValue: true })
  })

  it('converts undefined', () => {
    const cellValue = cellToExcel(null)
    expect(cellValue).toEqual({ type: ExcelCellValueType.empty })
  })

  it('converts null', () => {
    const cellValue = cellToExcel(undefined)
    expect(cellValue).toEqual({ type: ExcelCellValueType.empty })
  })

  it('converts a formatted number', () => {
    const cellValue = cellToExcel(new FormattedNumber(123, '0.00'))
    expect(cellValue).toEqual({ type: ExcelCellValueType.formattedNumber, basicValue: 123, numberFormat: '0.00' })
  })

  it('converts a moment', () => {
    const date = moment('2020-01-01')
    const cellValue = cellToExcel(date)
    expect(cellValue).toEqual({ type: ExcelCellValueType.double, basicValue: date.toOADate() })
  })

  it('returns the raw Excel CellValue if the type is unknown', () => {
    const cellValue = cellToExcel({ type: ExcelCellValueType.entity, text: 'hello!' })
    expect(cellValue).toEqual({ type: ExcelCellValueType.entity, text: 'hello!' })
  })

  it('converts anything else to a string', () => {
    const cellValue = cellToExcel({ foo: 'bar' } as any)
    expect(cellValue).toEqual({ type: ExcelCellValueType.string, basicValue: '[object Object]' })
  })
})
