import moment, { type Moment } from 'moment-msdate'
import { ExcelCellValueType } from 'vue-excel/types'
export type TableRecord = Record<string, any>

export type TableRow = {
  rowIndex: number
  range: Excel.Range
  record: TableRecord
}

export class FormattedNumber {
  number: number
  format: string

  constructor(number: number, format: string) {
    this.number = number
    this.format = format
  }

  toMoment(offset?: number | string): Moment {
    return moment.fromOADate(this.number, offset)
  }

  toNumber(): number {
    return Number(this.number)
  }
}

export type CellValue = string | number | boolean | null | undefined | Excel.CellValue | FormattedNumber | Moment

export class BaseTransformer {
  type!: ExcelCellValueType

  appliesToExcel(cellValue: CellValue): boolean {
    throw 'Not implemented!'
  }

  appliesFromExcel(cellValue: Excel.CellValue): boolean {
    return cellValue.type == this.type
  }

  toExcel(cellValue: CellValue): Excel.CellValue {
    return <Excel.CellValue>{ type: this.type, basicValue: cellValue }
  }

  fromExcel(cellValue: Excel.CellValue): CellValue {
    return cellValue.basicValue
  }
}

class BooleanTransformer extends BaseTransformer {
  type = ExcelCellValueType.boolean

  appliesToExcel(cellValue: CellValue): boolean {
    return typeof cellValue == 'boolean'
  }
}

class DoubleTransformer extends BaseTransformer {
  type = ExcelCellValueType.double

  appliesToExcel(cellValue: CellValue): boolean {
    return typeof cellValue == 'number'
  }
}

class StringTransformer extends BaseTransformer {
  type = ExcelCellValueType.string

  appliesToExcel(cellValue: CellValue): boolean {
    return typeof cellValue == 'string'
  }
}

class EmptyTransformer extends BaseTransformer {
  type = ExcelCellValueType.empty

  appliesToExcel(cellValue: CellValue): boolean {
    return cellValue == null || cellValue == undefined
  }

  fromExcel(cellValue: Excel.EmptyCellValue): CellValue {
    return null
  }

  toExcel(cellValue: null | undefined): Excel.CellValue {
    return <Excel.EmptyCellValue>{ type: this.type }
  }
}

class FormattedNumberTransformer extends BaseTransformer {
  type = ExcelCellValueType.formattedNumber

  appliesToExcel(cellValue: CellValue): boolean {
    return cellValue instanceof FormattedNumber
  }

  toExcel(cellValue: FormattedNumber): Excel.CellValue {
    return <Excel.FormattedNumberCellValue>{
      type: this.type,
      basicValue: cellValue.number,
      numberFormat: cellValue.format
    }
  }

  fromExcel(cellValue: Excel.FormattedNumberCellValue): FormattedNumber {
    return new FormattedNumber(cellValue.basicValue, cellValue.numberFormat)
  }
}

class MomentTransformer extends BaseTransformer {
  appliesToExcel(cellValue: CellValue): boolean {
    return cellValue instanceof moment
  }

  appliesFromExcel(cellValue: Excel.CellValue): boolean {
    return false
  }

  toExcel(cellValue: Moment): Excel.CellValue {
    return <Excel.DoubleCellValue>{ type: ExcelCellValueType.double, basicValue: cellValue.toOADate() }
  }
}

export const TransformerRegistry: { transformers: BaseTransformer[] } = {
  transformers: [
    new BooleanTransformer(),
    new DoubleTransformer(),
    new StringTransformer(),
    new EmptyTransformer(),
    new FormattedNumberTransformer(),
    new MomentTransformer()
  ]
}

export function cellToExcel(cellValue: CellValue): Excel.CellValue {
  const transformer = TransformerRegistry.transformers.find((t) => t.appliesToExcel(cellValue))

  if (transformer) {
    return transformer.toExcel(cellValue)
  } else {
    let excelCellValue = cellValue as Excel.CellValue
    if (Object.values(ExcelCellValueType).includes(<ExcelCellValueType>excelCellValue.type)) {
      return excelCellValue
    }
    return <Excel.StringCellValue>{ type: ExcelCellValueType.string, basicValue: cellValue?.toString() }
  }
}

export function cellFromExcel(cellValue: Excel.CellValue): CellValue {
  const transformer = TransformerRegistry.transformers.find((t) => t.appliesFromExcel(cellValue))

  if (transformer) {
    console.dir(transformer)
    return transformer.fromExcel(cellValue)
  } else {
    return cellValue
  }
}

export function rangeValuesToExcel(values: CellValue[][]): Excel.CellValue[][] {
  return values.map((rowValues: CellValue[]) => {
    return rowValues.map(cellToExcel)
  })
}

export function rangeValuesFromExcel(values: Excel.CellValue[][]): CellValue[][] {
  return values.map((rowValues: Excel.CellValue[]) => {
    return rowValues.map(cellFromExcel)
  })
}
