import xmlbuilder from 'xmlbuilder';
import { i2a } from '../util/i2a';
import { Workbook } from '../workbook';

export class CalcChain {
  book: Workbook;
  cache: Record<number, string[]> = {};

  constructor(book: Workbook) {
    this.book = book;
  }

  add_ref(idx: number, col: number, row: number) {
    const num = idx + 1;
    if (!this.cache.hasOwnProperty(num)) this.cache[num] = [];
    this.cache[num].push(i2a(col) + row);
  }
  toxml() {
    const cc = xmlbuilder.create('calcChain', {
      version: '1.0',
      encoding: 'UTF-8',
      standalone: true,
    });
    cc.att(
      'xmlns',
      'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    );

    for (const [key, val] of Object.entries(this.cache)) {
      for (const el of val) {
        cc.ele('c', { r: '' + el, i: '' + key });
      }
    }

    return cc.end({ pretty: false });
  }
}
