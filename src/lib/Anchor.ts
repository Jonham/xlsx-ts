import { TODO } from '../types';

export class Anchor {
  worksheet: TODO;

  nativeCol = 0;
  nativeColOff = 0;
  nativeRow = 0;
  nativeRowOff = 0;

  // TODO
  // colWidth: number = 0
  // rowHeight: number = 0

  constructor(worksheet: TODO, address?: TODO, offset?: TODO = 0) {
    this.worksheet = worksheet;

    if (!address) {
      this.nativeCol = 0;
      this.nativeColOff = 0;
      this.nativeRow = 0;
      this.nativeRowOff = 0;
    } else if (typeof address == 'string') {
      const decoded = colCache.decodeAddress(address);
      this.nativeCol = decoded.col + offset;
      this.nativeColOff = 0;
      this.nativeRow = decoded.row + offset;
      this.nativeRowOff = 0;
    } else if (address.nativeCol != undefined) {
      this.nativeCol = address.nativeCol || 0;
      this.nativeColOff = address.nativeColOff || 0;
      this.nativeRow = address.nativeRow || 0;
      this.nativeRowOff = address.nativeRowOff || 0;
    } else if (address.col != undefined) {
      this.col = address.col + offset;
      this.row = address.row + offset;
    } else {
      this.nativeCol = 0;
      this.nativeColOff = 0;
      this.nativeRow = 0;
      this.nativeRowOff = 0;
    }
    return;
  }

  // col getter setter
  get col(): number {
    return (
      this.nativeCol +
      Math.min(this.colWidth - 1, this.nativeColOff) / this.colWidth
    );
  }
  set col(v: number) {
    this.nativeCol = Math.floor(v);
    this.nativeColOff = Math.floor((v - this.nativeCol) * this.colWidth);
    return;
  }
  // col
  // enumerable: true
  // configurable: true

  // row getter setter
  get row(): number {
    return (
      this.nativeRow +
      Math.min(this.rowHeight - 1, this.nativeRowOff) / this.rowHeight
    );
  }
  set row(v: number) {
    this.nativeRow = Math.floor(v);
    this.nativeRowOff = Math.floor((v - this.nativeRow) * this.rowHeight);
    return;
  }
  // enumerable: true
  // configurable: true

  // colWidth getter setter
  get colWidth(): number {
    return 0;
    // if (this.worksheet && this.worksheet.width(this.nativeCol, this.nativeCol + 1)) then
    //  and @worksheet.getColumn(@nativeCol + 1).isCustomWidth then
    // Math.floor(@worksheet.getColumn(@nativeCol + 1).width * 10000) else 640000
  }
  // enumerable: true
  // configurable: true

  // rowHeight getter setter
  get rowHeight(): number {
    if (
      this.worksheet &&
      this.worksheet.getRow(this.nativeRow + 1) &&
      this.worksheet.getRow(this.nativeRow + 1).height
    ) {
      return Math.floor(
        this.worksheet.getRow(this.nativeRow + 1).height * 10000,
      );
    }
    return 180000;
  }
  // enumerable: true
  // configurable: true

  // model getter setter
  get model() {
    return {
      nativeCol: this.nativeCol,
      nativeColOff: this.nativeColOff,
      nativeRow: this.nativeRow,
      nativeRowOff: this.nativeRowOff,
    };
  }
  set model(value: TODO) {
    this.nativeCol = value.nativeCol;
    this.nativeColOff = value.nativeColOff;
    this.nativeRow = value.nativeRow;
    this.nativeRowOff = value.nativeRowOff;
    return;
  }
  // enumerable: true
  // configurable: true

  asInstance(model?: Anchor | null) {
    return model instanceof Anchor || model == null ? model : new Anchor(model);
  }

  toxml(xml: TODO) {
    const wb = xml.create('workbook', {
      version: '1.0',
      encoding: 'UTF-8',
      standalone: true,
    });
    wb.ele('from').ele('workbookView', {
      xWindow: '0',
      yWindow: '90',
      windowWidth: '19200',
      windowHeight: '11640',
    });

    wb.att(
      'xmlns',
      'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    );
    wb.att(
      'xmlns:r',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    );

    wb.ele('fileVersion', {
      appName: 'xl',
      lastEdited: '4',
      lowestEdited: '4',
      rupBuild: '4505',
    });
    wb.ele('workbookPr', { filterPrivacy: '1', defaultThemeVersion: '124226' });
    wb.ele('bookViews').ele('workbookView', {
      xWindow: '0',
      yWindow: '90',
      windowWidth: '19200',
      windowHeight: '11640',
    });
  }
}
