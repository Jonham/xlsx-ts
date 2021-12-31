import { Workbook } from '.';
import { TODO, _Image } from '../types';

export type PageMargin = {
  /** 默认值: '0.7' */
  left: string;
  /** 默认值: '0.7' */
  right: string;
  /** 默认值: '0.75' */
  top: string;
  /** 默认值: '0.75' */
  bottom: string;
  /** 默认值: '0.3' */
  header: string;
  /** 默认值: '0.3' */
  footer: string;
};
export function getDefaultPageMargin(): PageMargin {
  return {
    left: '0.7',
    right: '0.7',
    top: '0.75',
    bottom: '0.75',
    header: '0.3',
    footer: '0.3',
  };
}

export type SheetData = any;
export type CellMerge = any;
export type ColumnWidth = any;
export type RowHeight = any;
export type CellStyle = any;
export type Formula = any;

export class Sheet {
  name: string;
  book: Workbook;
  data: Record<string, SheetData>;
  pageMargins: PageMargin;
  merges: CellMerge[];
  colWidths: ColumnWidth[];
  rowHeights: Record<string, RowHeight>;
  styles: Record<string, CellStyle>;
  formulas: Formula[];
  images: _Image[];
  // TODO
  wsRels: any[] = [];

  constructor(
    book: Workbook,
    name: string,
    colCount: number,
    rowCount: number,
  ) {
    this.book = book;
    this.name = name;
    this.data = {};
    for (let i = 1; i <= rowCount; i++) {
      this.data[i] = [];
      for (let j = 1; j <= colCount; j++) {
        this.data[i][j] = { v: 0 };
      }
    }

    this.merges = [];
    this.colWidths = [];
    this.rowHeights = {};
    this.styles = {};
    this.formulas = [];
    this.pageMargins = getDefaultPageMargin();
    this.images = [];
  }

  /**
    validates exclusivity between filling base64, filename, buffer properties.
    validates extension is among supported types.
    concurrency this is a critical path add semaphor, only one image can be added at the time.
    there's a risk of adding image in parallel and returing diferent id between push and returning.
    exceljs also contains same risk, despite of collecting id before.
   */
  addImage (image?: TODO) {
    // if (!image || !image.range || !image.base64 || !image.extension)
    //   throw Error('please verify your image format')


    // // tries to decode range
    // if ((typeof image.range != 'string') || !/\w+\d+:\w+\d/i.test(image.range))
    //   throw Error('Please provide range parameter like `B2:F6`.')

    // const decoded = colCache.decode(image.range);
    // this.range = {
    //   from: new Anchor(this.worksheet, decoded.tl, -1),
    //   to: new Anchor(this.worksheet, decoded.br, 0),
    //   editAs: 'oneCell',
    // }

    const id = this.book.medias.length + 1
    // const imageToAdd = new Image(id, image.extension, image.base64, this.range, image.options||{})
    // const media = this.book._addMediaFromImage(imageToAdd)
    // // drawingId = this.book._addDrawingFromImage(imageToAdd)
    // // wsDwRelId = this.sheet._addDrawingFromImage(imageToAdd)
    // console.log(imageToAdd)
    // this.images.push(imageToAdd)

    return id
  }

  // TODO
  toxml(): string {
    return '';
  }
}
