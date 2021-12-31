import { Workbook } from '.'
import { _Image } from '../types'

export type PageMargin = {
  /** 默认值: '0.7' */
  left: string
  /** 默认值: '0.7' */
  right: string
  /** 默认值: '0.75' */
  top: string
  /** 默认值: '0.75' */
  bottom: string
  /** 默认值: '0.3' */
  header: string
  /** 默认值: '0.3' */
  footer: string
}
export function getDefaultPageMargin(): PageMargin {
  return {
    left: '0.7',
    right: '0.7',
    top: '0.75',
    bottom: '0.75',
    header: '0.3',
    footer: '0.3',
  }
}

export type SheetData = any
export type CellMerge = any
export type ColumnWidth = any
export type RowHeight = any
export type CellStyle = any
export type Formula = any

export class Sheet {
  name: string
  data: Record<string, SheetData>
  pageMargins: PageMargin
  merges: CellMerge[]
  colWidths: ColumnWidth[]
  rowHeights: Record<string, RowHeight>
  styles: Record<string, CellStyle>
  formulas: Formula[]
  images: _Image[]
  // TODO
  wsRels: any[] = []

  constructor(
    book: Workbook,
    name: string,
    colCount: number,
    rowCount: number
  ) {
    this.name = name
    this.data = {}
    for (let i = 1; i <= rowCount; i++) {
      this.data[i] = []
      for (let j = 1; j <= colCount; j++) {
        this.data[i][j] = { v: 0 }
      }
    }

    this.merges = []
    this.colWidths = []
    this.rowHeights = {}
    this.styles = {}
    this.formulas = []
    this.images = []
    this.pageMargins = getDefaultPageMargin()
  }

  // TODO
  toxml (): string {
    return ''
  }
}
