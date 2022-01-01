import { Workbook } from '.';
import { Image, _range } from '../lib/Image';
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
  images: Image[];
  // TODO
  wsRels: any[] = [];
  // TODO
  range?: _range;
  // TODO
  worksheet?: TODO;

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
  addImage(image?: Image) {
    if (!image || !image.range || !image.base64 || !image.extension)
      throw Error('please verify your image format');

    // tries to decode range
    if (typeof image.range != 'string' || !/\w+\d+:\w+\d/i.test(image.range))
      throw Error('Please provide range parameter like `B2:F6`.');

    const decoded = colCache.decode(image.range);
    this.range = {
      from: new Anchor(this.worksheet, decoded.tl, -1),
      to: new Anchor(this.worksheet, decoded.br, 0),
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

  set (col: number, row: number, str: string | Date) {
    // if (arguments.length==1 && col && typeof col == 'object')
    //   cells = col
    //   for c,col of cells
    //     for r,cell of col
    //       this.set(c,r, cell)


    // else if str instanceof Date
    //   @set col, row, JSDateToExcel str
    //   # for some reason the number format doesn't apply if the fill is not also set. BUG? Mystery?
    //   @fill col, row,
    //     type: "solid",
    //     fgColor: "FFFFFF"
    //   @numberFormat col, row, 'd-mmm'
    // else if typeof str == 'object'
    //   for key of str
    //     @[key] col, row, str[key]
    // else if  typeof str == 'string'
    //   if str != null and str != ''
    //     @data[row][col].v = @book.ss.str2id('' + str)
    //   return @data[row][col].dataType = 'string'
    // else if typeof str == 'number'
    //   @data[row][col].v = str
    //   return @data[row][col].dataType = 'number'
    // else
    //   @data[row][col].v = str
    // return
  }
  

  // formula: (col, row, str) ->
  //   if (typeof str == 'string')
  //     @formulas = @formulas || []
  //     @formulas[row] = @formulas[row] || []
  //     sheet_idx = i for sheet, i in @book.sheets when sheet.name == @name
  //     @book.cc.add_ref(sheet_idx, col, row)
  //     @formulas[row][col] = str

  // merge: (from_cell, to_cell) ->
  //   @merges.push({from: from_cell, to: to_cell})

  // width: (col, wd) ->
  //   @col_wd.push({c: col, cw: wd})

  // getColWidth: (col) ->
  //   for _col in @col_wd
  //     if _col.c == col
  //       return Math.floor(_col.cw * 10000)
  //   return 640000

  // height: (row, ht) ->
  //   @row_ht[row] = ht

  // font: (col, row, font_s)->
  //   @styles['font_' + col + '_' + row] = @book.st.font2id(font_s)

  // fill: (col, row, fill_s)->
  //   @styles['fill_' + col + '_' + row] = @book.st.fill2id(fill_s)

  // border: (col, row, bder_s)->
  //   @styles['bder_' + col + '_' + row] = @book.st.bder2id(bder_s)

  // numberFormat: (col, row, numfmt_s)->
  //   @styles['numfmt_' + col + '_' + row] = @book.st.numfmt2id(numfmt_s)

  // align: (col, row, align_s)->
  //   @styles['algn_' + col + '_' + row] = align_s

  // valign: (col, row, valign_s)->
  //   @styles['valgn_' + col + '_' + row] = valign_s

  // rotate: (col, row, textRotation)->
  //   @styles['rotate_' + col + '_' + row] = textRotation

  // wrap: (col, row, wrap_s)->
  //   @styles['wrap_' + col + '_' + row] = wrap_s

  // autoFilter: (filter_s) ->
  //   @autofilter = if typeof filter_s == 'string' then filter_s else @getRange()

  // _sheetViews: {
  //   workbookViewId: '0'
  // }

  // _sheetViewsPane: {

  // }

  // _pageSetup: {
  //   paperSize: '9',
  //   orientation: 'portrait',
  //   horizontalDpi: '200',
  //   verticalDpi: '200'
  //   }

  // sheetViews: (obj) ->
  //   for key, val of obj
  //     if (typeof this[key] == 'function')
  //       this[key](obj[key])
  //     else
  //       @_sheetViews[key] = val

  // split: (ncols, nrows, state, activePane, topLeftCell) ->
  //   state = state || "frozen"
  //   activePane = activePane || "bottomRight"
  //   topLeftCell = topLeftCell || (tool.i2a((ncols || 0) + 1) + ((nrows || 0) + 1))
  //   if (ncols)
  //     @_sheetViewsPane.xSplit = '' + ncols
  //   if (nrows)
  //     @_sheetViewsPane.ySplit = '' + nrows
  //   if (state)
  //     @_sheetViewsPane.state = state
  //   if (activePane)
  //     @_sheetViewsPane.activePane = activePane
  //   if (topLeftCell)
  //     @_sheetViewsPane.topLeftCell = topLeftCell

  // printBreakRows: (arr) ->
  //   @_rowBreaks = arr

  // printBreakColumns: (arr) ->
  //   @_colBreaks = arr


  // printRepeatRows: (start, end) ->
  //   if Array.isArray(start)
  //     @_repeatRows = {start: start[0], end: start[1]}
  //   else
  //     @_repeatRows = {start, end}

  // printRepeatColumns: (start, end) ->
  //   if Array.isArray(start)
  //     @_repeatCols = {start: start[0], end: start[1]}
  //   else @_repeatCols =  {start, end}

  // pageSetup: (obj) ->
  //   for key, val of obj
  //     @_pageSetup[key] = val

  // pageMargins: (obj) ->
  //   for key, val of obj
  //     @_pageMargins[key] = val


  // style_id: (col, row) ->
  //   inx = '_' + col + '_' + row
  //   style = {
  //     numfmt_id: @styles['numfmt' + inx],
  //     font_id: @styles['font' + inx],
  //     fill_id: @styles['fill' + inx],
  //     bder_id: @styles['bder' + inx],
  //     align: @styles['algn' + inx],
  //     valign: @styles['valgn' + inx],
  //     rotate: @styles['rotate' + inx],
  //     wrap: @styles['wrap' + inx]
  //   }
  //   id = @book.st.style2id(style)
  //   return  id

  // getRange: () ->
  //   return '$A$1:$' + tool.i2a(@cols) + '$' + @rows


  // TODO
  toxml(): string {
    return '';
    // ws = xml.create('worksheet', {version: '1.0', encoding: 'UTF-8', standalone: true})
    // ws.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    // ws.att('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    // #    ws.att('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')

    // #    ws.att('mc:Ignorable', "x14ac xr xr2 xr3")
    // #    ws.att('xmlns:x14ac', "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
    // #    ws.att('xmlns:xr', "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
    // #    ws.att('xmlns:xr2', "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
    // #    ws.att('xmlns:xr3', "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3")

    // ws.ele('dimension', {ref: 'A1'})

    // ws.ele('sheetViews').ele('sheetView', @_sheetViews).ele('pane', @_sheetViewsPane)


    // ws.ele('sheetFormatPr', {defaultRowHeight: '13.5'})
    // if @col_wd.length > 0
    //   cols = ws.ele('cols')
    //   for cw in @col_wd
    //     cols.ele('col', {min: '' + cw.c, max: '' + cw.c, width: cw.cw, customWidth: '1'})
    // sd = ws.ele('sheetData')
    // for i in [1..@rows]
    //   r = sd.ele('row', {r: '' + i, spans: '1:' + @cols})
    //   ht = @row_ht[i]
    //   if ht
    //     r.att('ht', ht)
    //     r.att('customHeight', '1')
    //   for j in [1..@cols]
    //     ix = @data[i][j]
    //     sid = @style_id(j, i)
    //     if (ix.v isnt null and ix.v isnt undefined) or (sid isnt 1)
    //       c = r.ele('c', {r: '' + tool.i2a(j) + i})
    //       c.att('s', '' + (sid - 1)) if sid isnt 1
    //       if (@formulas[i] && @formulas[i][j])
    //         c.ele('f', '' + @formulas[i][j])
    //         c.ele('v')
    //       else if ix.dataType == 'string'
    //         c.att('t', 's')
    //         c.ele('v', '' + (ix.v - 1))
    //       else if ix.dataType == 'number'
    //         c.ele 'v', '' + ix.v



    // if @merges.length > 0
    //   mc = ws.ele('mergeCells', {count: @merges.length})
    //   for m in @merges
    //     mc.ele('mergeCell', {ref: ('' + tool.i2a(m.from.col) + m.from.row + ':' + tool.i2a(m.to.col) + m.to.row)})
    // if typeof @autofilter == 'string'
    //   ws.ele('autoFilter', {ref: @autofilter})
    // ws.ele('phoneticPr', {fontId: '1', type: 'noConversion'})

    // ws.ele('pageMargins', @_pageMargins)
    // ws.ele('pageSetup', @_pageSetup)

    // if @_rowBreaks && @_rowBreaks.length
    //   cb = ws.ele('rowBreaks', {count: @_rowBreaks.length, manualBreakCount: @_rowBreaks.length})
    //   for i in @_rowBreaks
    //     cb.ele('brk', { id: i, man: '1'})

    // if @_colBreaks && @_colBreaks.length
    //   cb = ws.ele('colBreaks', {count: @_colBreaks.length, manualBreakCount: @_colBreaks.length})
    //   for i in @_colBreaks
    //     cb.ele('brk', { id: i, man: '1'})

    // for wsRel in @wsRels
    //   ws.ele('drawing', {'r:id': wsRel.id})

    // ws.end({pretty: false})
  }
}
