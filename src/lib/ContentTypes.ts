import xmlbuilder from 'xmlbuilder';
import { Workbook } from '../workbook';

export class ContentTypes {
  book: Workbook;

  constructor(book: Workbook) {
    this.book = book;
  }

  // ??
  _getKnowImageTypes() {
    const imagesToAdd = [];
    // debugger;
    const imagesToAddDistinct: Record<string, string> = {};
    for (const sheet of this.book.sheets) {
      for (const image of sheet.images) {
        if (!imagesToAddDistinct[image.extension]) {
          imagesToAdd.push({
            Extension: image.extension,
            ContentType: image.contentType,
          });
        }
      }
    }
    return imagesToAdd;
  }

  toxml(): string {
    const types = xmlbuilder.create('Types', {
      version: '1.0',
      encoding: 'UTF-8',
      standalone: true,
    });
    types.att(
      'xmlns',
      'http://schemas.openxmlformats.org/package/2006/content-types',
    );

    // TODO remove fixed types here
    types.ele('Default', { Extension: 'png', ContentType: 'image/png' });
    types.ele('Default', { Extension: 'svg', ContentType: 'image/svg+xml' });
    types.ele('Default', {
      Extension: 'rels',
      ContentType: 'application/vnd.openxmlformats-package.relationships+xml',
    });
    types.ele('Default', { Extension: 'xml', ContentType: 'application/xml' });

    types.ele('Override', {
      PartName: '/xl/workbook.xml',
      ContentType:
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
    });
    for (let i = 1; i <= this.book.sheets.length; i++) {
      types.ele('Override', {
        PartName: '/xl/worksheets/sheet' + i + '.xml',
        ContentType:
          'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
      });
    }

    types.ele('Override', {
      PartName: '/xl/theme/theme1.xml',
      ContentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
    });
    types.ele('Override', {
      PartName: '/xl/calcChain.xml',
      ContentType:
        'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml',
    });
    types.ele('Override', {
      PartName: '/xl/styles.xml',
      ContentType:
        'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
    });
    for (const sheet of this.book.sheets) {
      for (const image of sheet.images) {
        types.ele('Override', {
          PartName: '/xl/drawings/drawing' + image.id + '.xml',
          ContentType:
            'application/vnd.openxmlformats-officedocument.drawing+xml',
        });
      }
    }

    types.ele('Override', {
      PartName: '/docProps/core.xml',
      ContentType: 'application/vnd.openxmlformats-package.core-properties+xml',
    });
    types.ele('Override', {
      PartName: '/docProps/app.xml',
      ContentType:
        'application/vnd.openxmlformats-officedocument.extended-properties+xml',
    });

    //    if Object.getOwnPropertyNames(@book.cc.cache).length > 0
    //      types.ele('Override', {
    //        PartName: '/xl/calcChain.xml',
    //        ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml'
    //      });
    // for knowImageType in @_getKnowImageTypes
    // 	types.ele('Default', knowImageType)
    types.ele('Override', {
      PartName: '/xl/sharedStrings.xml',
      ContentType:
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
    });

    return types.end({ pretty: false });
  }
}
