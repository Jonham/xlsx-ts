import xmlbuilder from 'xmlbuilder';
import { Workbook } from '.';
import { Anchor, AnchorRange } from '../lib/Anchor';
import { Image } from '../lib/Image';
import { Border, Fill, FontDef, StyleDef, TODO } from '../types';
import { colCache, _Unknown } from '../util/colCache';
import { i2a } from '../util/i2a';
import { JSDateToExcel } from '../util/JSDateToExcel';
import { getDefaultPageMargin, PageMargin } from '../util/pageMargin';

type ColWd = {
  c: string;
  cw: number;
};

// TODO
// type RowBreak = {};
type RowBreak = number;
type ColBreak = {};

type PageSetup = {
  paperSize: string;
  orientation: string;
  horizontalDpi: string;
  verticalDpi: string;
};

type Cell = any;
export type CellMerge = {
  from: Cell;
  to: Cell;
};

type SheetViewPane = {
  xSplit: TODO;
  ySplit: TODO;
  state: TODO;
  activePane: TODO;
  topLeftCell: TODO;
};

export type SheetData = any;
export type ColumnWidth = any;
export type RowHeight = any;
// export type CellStyle = any;
export type Formula = any;

export class Sheet {
  name: string;
  book: Workbook;
  data: Record<string, SheetData>;
  // pageMargins: PageMargin;
  _pageMargins: PageMargin;
  merges: CellMerge[];
  colWidths: ColumnWidth[];
  rowHeights: Record<string, RowHeight>;
  /** TODO change to Map number and Map string */
  styles: Record<string, number | string>;

  formulas: Formula[];
  images: Image[];
  // TODO
  wsRels: any[] = [];
  // TODO
  range?: AnchorRange;
  // TODO
  worksheet?: Sheet;
  // TODO
  autofilter?: string;

  _repeatRows?: {
    start: number;
    end: number;
  };
  _repeatCols?: {
    start: number;
    end: number;
  };

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
    this._pageMargins = getDefaultPageMargin();
    this.images = [];
  }

  /**
    validates exclusivity between filling base64, filename, buffer properties.
    validates extension is among supported types.
    concurrency this is a critical path add semaphor, only one image can be added at the time.
    there's a risk of adding image in parallel and returing diferent id between push and returning.
    exceljs also contains same risk, despite of collecting id before.
   */
  addImage(image?: Image) {
    if (!image || !image.range || !image.base64 || !image.extension)
      throw Error('please verify your image format');

    // tries to decode range
    if (typeof image.range != 'string' || !/\w+\d+:\w+\d/i.test(image.range))
      throw Error('Please provide range parameter like `B2:F6`.');

    const decoded = colCache.decode(image.range) as _Unknown;
    this.range = {
      from: new Anchor(this.worksheet as Sheet, decoded.tl, -1),
      to: new Anchor(this.worksheet as Sheet, decoded.br, 0),
      editAs: 'oneCell',
    };

    const id = this.book.medias.length + 1;
    const imageToAdd = new Image(
      id,
      image.extension,
      image.base64,
      this.range,
      image.options || {},
    );
    const media = this.book._addMediaFromImage(imageToAdd);
    // const drawingId = this.book._addDrawingFromImage(imageToAdd)
    // wsDwRelId = this.sheet._addDrawingFromImage(imageToAdd)
    console.log(imageToAdd);
    this.images.push(imageToAdd);

    return id;
  }

  getImage(id: number) {
    return this.images[id];
  }

  getImages() {
    return this.images;
  }

  removeImage(id: number) {
    this.images = this.images.filter((i) => i.id !== id);
  }

  /** old approach for adding background images.
    addBackgroundImage: (imageId) ->
      model = {
        type: 'background',
        imageId,
      }
      @_media.push(new Image(this, model))

    getBackgroundImageId: ()->
      image = @_media.find(m => m.type == 'background')
      return image && image.imageId
  */

  set(col: any): void;
  set(col: number, row: number, str: any): void;
  set(...args: any[]): void {
    const [col, row, str] = args;
    // TODO
    if (args.length === 1 && col && typeof col == 'object') {
      const cells = col;
      for (const [c, col] of Object.entries(cells)) {
        for (const [r, cell] of Object.entries(col as any)) {
          this.set(parseInt(c), parseInt(r), cell as any);
        }
      }
    } else if (str instanceof Date) {
      this.set(col, row, JSDateToExcel(str));
      // TODO
      // for some reason the number format doesn't apply if the fill is not also set. BUG? Mystery?
      this.fill(col, row, {
        type: 'solid',
        fgColor: 'FFFFFF',
      });
      this.numberFormat(col, row, 'd-mmm');
    } else if (typeof str === 'object') {
      for (const [key, value] of Object.entries(str)) {
        // TODO
        // @ts-ignore
        this[key](col, row, value);
      }
    } else if (typeof str == 'string') {
      // if (str != null && str !== '') {} // ??
      if (str !== '') {
        this.data[row][col].v = this.book.sharedStrings.str2id('' + str);
      }
      this.data[row][col].dataType = 'string'; // ?? return
      return;
    } else if (typeof str == 'number') {
      this.data[row][col].v = str;
      this.data[row][col].dataType = 'number'; // ?? return
      return;
    } else {
      this.data[row][col].v = str;
    }
    return;
  }

  formula(col: number, row: number, str: string) {
    if (typeof str == 'string') {
      this.formulas = this.formulas || [];
      this.formulas[row] = this.formulas[row] || [];
      // sheet_idx = i for sheet, i in this.book.sheets when sheet.name == this.name
      const sheet_idx = this.book.sheets.findIndex(
        (sheet) => sheet.name === this.name,
      );
      this.book.calcChain.add_ref(sheet_idx, col, row);
      this.formulas[row][col] = str;
    }
  }

  merge(from_cell: Cell, to_cell: Cell) {
    this.merges.push({ from: from_cell, to: to_cell });
  }

  col_wd: ColWd[] = [];
  row_ht: number[] = [];

  width(col: string, wd: number) {
    return this.col_wd.push({ c: col, cw: wd });
  }

  getColWidth(col: string) {
    for (const _col of this.col_wd) {
      if (_col.c == col) return Math.floor(_col.cw * 10000);
    }
    return 640000;
  }

  height(row: number, ht: number) {
    this.row_ht[row] = ht;
  }

  font(col: number, row: number, font_s: Partial<FontDef>) {
    return (this.styles['font_' + col + '_' + row] =
      this.book.style.font2id(font_s));
  }

  fill(col: number, row: number, fill_s: Partial<Fill>) {
    return (this.styles['fill_' + col + '_' + row] =
      this.book.style.fill2id(fill_s));
  }

  border(col: number, row: number, bder_s: Partial<Border>) {
    return (this.styles['bder_' + col + '_' + row] =
      this.book.style.bder2id(bder_s));
  }

  numberFormat(col: number, row: number, numfmt_s: any) {
    this.styles['numfmt_' + col + '_' + row] =
      this.book.style.numfmt2id(numfmt_s);
  }

  // TODO
  align(col: number, row: number, align_s: any) {
    return (this.styles['algn_' + col + '_' + row] = align_s);
  }
  // TODO
  valign(col: number, row: number, valign_s: any) {
    return (this.styles['valgn_' + col + '_' + row] = valign_s);
  }
  // TODO
  rotate(col: number, row: number, textRotation: any) {
    return (this.styles['rotate_' + col + '_' + row] = textRotation);
  }
  // TODO
  wrap(col: number, row: number, wrap_s: any) {
    return (this.styles['wrap_' + col + '_' + row] = wrap_s);
  }
  // TODO
  autoFilter(filter_s: string) {
    return (this.autofilter =
      typeof filter_s === 'string' ? filter_s : this.getRange());
  }

  _sheetViews = {
    workbookViewId: '0',
  };

  _sheetViewsPane: Partial<SheetViewPane> = {};

  _pageSetup: PageSetup = {
    paperSize: '9',
    orientation: 'portrait',
    horizontalDpi: '200',
    verticalDpi: '200',
  };

  sheetViews(obj: Record<'workbookViewId' & keyof Sheet, any>) {
    for (const [key, val] of Object.entries(obj)) {
      const k = key as 'workbookViewId' & keyof Sheet;
      const fn = this[k];
      if (typeof fn === 'function') {
        // @ts-ignore
        this[k](val);
      } else {
        // @ts-ignore
        this._sheetViews[k] = val;
      }
    }
  }

  split(
    ncols: number,
    nrows: number,
    state = 'frozen',
    activePane = 'bottomRight',
    _topLeftCell?: TODO,
  ) {
    // const state = state || "frozen"
    // activePane = activePane || "bottomRight"
    const topLeftCell =
      _topLeftCell || i2a((ncols || 0) + 1) + ((nrows || 0) + 1);
    if (ncols) this._sheetViewsPane.xSplit = '' + ncols;
    if (nrows) this._sheetViewsPane.ySplit = '' + nrows;
    if (state) this._sheetViewsPane.state = state;
    if (activePane) this._sheetViewsPane.activePane = activePane;
    if (topLeftCell) this._sheetViewsPane.topLeftCell = topLeftCell;
  }

  _rowBreaks: RowBreak[] = [];
  _colBreaks: ColBreak[] = [];
  printBreakRows(arr: RowBreak[]) {
    this._rowBreaks = arr;
  }
  printBreakColumns(arr: ColBreak[]) {
    this._colBreaks = arr;
  }

  printRepeatRows(start: number | [number, number], end: number) {
    if (Array.isArray(start)) {
      this._repeatRows = { start: start[0], end: start[1] };
    } else {
      this._repeatRows = { start, end };
    }
  }

  printRepeatColumns(start: number | [number, number], end: number) {
    if (Array.isArray(start)) {
      this._repeatCols = { start: start[0], end: start[1] };
    } else {
      this._repeatCols = { start, end };
    }
  }

  pageSetup(obj: any & TODO) {
    for (const [key, val] of Object.entries(obj)) {
      this._pageSetup[key as keyof PageSetup] = val as string;
    }
  }

  pageMargins(obj: Partial<PageMargin>) {
    for (const [key, val] of Object.entries(obj)) {
      this._pageMargins[key as keyof PageMargin] = val as string;
    }
  }

  style_id(col: number, row: number) {
    const inx = '_' + col + '_' + row;
    const style: StyleDef = {
      numfmt_id: this.styles['numfmt' + inx] as number,
      font_id: this.styles['font' + inx] as number,
      fill_id: this.styles['fill' + inx] as number,
      bder_id: this.styles['bder' + inx] as number,

      // TODO
      align: this.styles['algn' + inx] as string,
      valign: this.styles['valgn' + inx] as string,
      rotate: this.styles['rotate' + inx] as string,
      wrap: this.styles['wrap' + inx] as string,
    };
    const id = this.book.style.style2id(style);
    return id;
  }

  // TODO
  cols: number = 0; // ??
  rows: number = 0; // ??

  getRange() {
    return '$A$1:$' + i2a(this.cols) + '$' + this.rows;
  }

  // TODO
  toxml(): string {
    // return '';
    const ws = xmlbuilder.create('worksheet', {
      version: '1.0',
      encoding: 'UTF-8',
      standalone: true,
    });
    ws.att(
      'xmlns',
      'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    );
    ws.att(
      'xmlns:r',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    );
    // ws.att('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
    // ws.att('mc:Ignorable', "x14ac xr xr2 xr3")
    // ws.att('xmlns:x14ac', "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
    // ws.att('xmlns:xr', "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
    // ws.att('xmlns:xr2', "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
    // ws.att('xmlns:xr3', "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3")

    ws.ele('dimension', { ref: 'A1' });

    ws.ele('sheetViews')
      .ele('sheetView', this._sheetViews)
      .ele('pane', this._sheetViewsPane);

    ws.ele('sheetFormatPr', { defaultRowHeight: '13.5' });
    if (this.col_wd.length > 0) {
      const cols = ws.ele('cols');
      for (const cw of this.col_wd) {
        cols.ele('col', {
          min: '' + cw.c,
          max: '' + cw.c,
          width: cw.cw,
          customWidth: '1',
        });
      }
    }

    const sd = ws.ele('sheetData');
    for (let i = 1; i <= this.rows; i++) {
      const r = sd.ele('row', { r: '' + i, spans: '1:' + this.cols });
      const ht = this.row_ht[i];
      if (ht) {
        r.att('ht', ht);
        r.att('customHeight', '1');
      }

      for (let j = 1; j <= this.cols; j++) {
        const ix = this.data[i][j];
        const sid = this.style_id(j, i);
        if ((ix.v !== null && ix.v !== undefined) || sid !== 1) {
          const c = r.ele('c', { r: '' + i2a(j) + i });
          if (sid !== 1) c.att('s', '' + (sid - 1));
          if (this.formulas[i] && this.formulas[i][j]) {
            c.ele('f', '' + this.formulas[i][j]);
            c.ele('v');
          } else if (ix.dataType == 'string') {
            c.att('t', 's');
            c.ele('v', '' + (ix.v - 1));
          } else if (ix.dataType == 'number') {
            c.ele('v', '' + ix.v);
          }
        }
      }
    }

    if (this.merges.length > 0) {
      const mc = ws.ele('mergeCells', { count: this.merges.length });
      for (const m of this.merges) {
        mc.ele('mergeCell', {
          ref:
            '' + i2a(m.from.col) + m.from.row + ':' + i2a(m.to.col) + m.to.row,
        });
      }
    }
    if (typeof this.autofilter == 'string') {
      ws.ele('autoFilter', { ref: this.autofilter });
    }
    ws.ele('phoneticPr', { fontId: '1', type: 'noConversion' });

    ws.ele('pageMargins', this._pageMargins);
    ws.ele('pageSetup', this._pageSetup);

    if (this._rowBreaks && this._rowBreaks.length) {
      const cb = ws.ele('rowBreaks', {
        count: this._rowBreaks.length,
        manualBreakCount: this._rowBreaks.length,
      });
      for (const i of this._rowBreaks) {
        cb.ele('brk', { id: i, man: '1' });
      }
    }

    if (this._colBreaks && this._colBreaks.length) {
      const cb = ws.ele('colBreaks', {
        count: this._colBreaks.length,
        manualBreakCount: this._colBreaks.length,
      });
      for (const i of this._colBreaks) {
        cb.ele('brk', { id: i, man: '1' });
      }
    }

    for (const wsRel of this.wsRels) ws.ele('drawing', { 'r:id': wsRel.id });

    return ws.end({ pretty: false });
  }

  // TODO
  getRow(rn: number): { height: number } {
    console.log(rn);
    console.error('Sheet.getRow() is NOT implement');
    return {
      height: 0,
    };
  }
}
