import xmlbuilder from 'xmlbuilder';
import { Workbook } from '../workbook';

export class DocPropsApp {
  book: Workbook;
  constructor(book: Workbook) {
    this.book = book;
  }

  toxml(): string {
    const props = xmlbuilder.create('Properties', {
      version: '1.0',
      encoding: 'UTF-8',
      standalone: true,
    });
    props.att(
      'xmlns',
      'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    );
    props.att(
      'xmlns:vt',
      'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
    );
    props.ele('Application', 'Microsoft Excel');
    props.ele('DocSecurity', '0');
    props.ele('ScaleCrop', 'false');

    let tmp = props
      .ele('HeadingPairs')
      .ele('vt:vector', { size: 2, baseType: 'variant' });
    tmp.ele('vt:variant').ele('vt:lpstr', 'Worksheets');
    tmp.ele('vt:variant').ele('vt:i4', '' + this.book.sheets.length);

    tmp = props
      .ele('TitlesOfParts')
      .ele('vt:vector', { size: this.book.sheets.length, baseType: 'lpstr' });
    for (let i = 1; i <= this.book.sheets.length; i++) {
      tmp.ele('vt:lpstr', this.book.sheets[i - 1].name);
    }
    props.ele('Company');
    props.ele('LinksUpToDate', 'false');
    props.ele('SharedDoc', 'false');
    props.ele('HyperlinksChanged', 'false');
    props.ele('AppVersion', '12.0000');
    return props.end({ pretty: false });
  }
}
