import xmlbuilder from 'xmlbuilder';
import { Workbook } from '../workbook';
import { i2a } from './i2a';

export class XlWorkbook {
  book: Workbook;
  constructor(book: Workbook) {
    this.book = book;
  }

  toxml() {
    const wb = xmlbuilder.create('workbook', {
      version: '1.0',
      encoding: 'UTF-8',
      standalone: true,
    });
    wb.att(
      'xmlns',
      'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    );
    wb.att(
      'xmlns:r',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    );
    wb.att(
      'xmlns:mc',
      'http://schemas.openxmlformats.org/markup-compatibility/2006',
    );
    wb.att('mc:Ignorable', 'x15 xr xr6 xr10 xr2');
    wb.att(
      'xmlns:x15',
      'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main',
    );
    wb.att(
      'xmlns:xr',
      'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
    );
    wb.att(
      'xmlns:xr6',
      'http://schemas.microsoft.com/office/spreadsheetml/2016/revision6',
    );
    wb.att(
      'xmlns:xr10',
      'http://schemas.microsoft.com/office/spreadsheetml/2016/revision10',
    );
    wb.att(
      'xmlns:xr2',
      'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2',
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

    const tmp = wb.ele('sheets');
    for (let i = 1; i <= this.book.sheets.length; i++) {
      tmp.ele('sheet', {
        name: this.book.sheets[i - 1].name,
        sheetId: '' + i,
        'r:id': 'rId' + i,
      });
    }

    const definedNames = wb.ele('definedNames'); // one entry per autofilter

    this.book.sheets.forEach((sheet, idx) => {
      if (sheet.autofilter) {
        definedNames
          .ele('definedName', {
            name: '_xlnm._FilterDatabase',
            hidden: '1',
            localSheetId: idx,
          })
          .raw("'" + sheet.name + "'!" + sheet.getRange());
      }

      if (sheet._repeatRows || sheet._repeatCols) {
        let range = '';
        if (sheet._repeatCols) {
          range +=
            "'" +
            sheet.name +
            "'!$" +
            i2a(sheet._repeatCols.start) +
            ':$' +
            i2a(sheet._repeatCols.end);
        }
        if (sheet._repeatRows) {
          range +=
            ",'" +
            sheet.name +
            "'!$" +
            sheet._repeatRows.start +
            ':$' +
            sheet._repeatRows.end;
        }

        definedNames
          .ele('definedName', {
            name: '_xlnm.Print_Titles',
            localSheetId: idx,
          })
          .raw(range);
      }
    });

    wb.ele('calcPr', { calcId: '124519' });
    return wb.end({ pretty: false });
  }
}
