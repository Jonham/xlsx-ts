import xmlbuilder from 'xmlbuilder';

export class SharedStrings {
  cache: Record<string, number>;
  arr: string[] = [];

  constructor() {
    this.cache = {};
    this.arr = [];
  }

  str2id(s: string): number {
    const id = this.cache[s];
    if (id) {
      return id;
    } else {
      this.arr.push(s);
      this.cache[s] = this.arr.length;
      return this.arr.length;
    }
  }

  toxml(): string {
    const sst = xmlbuilder.create('sst', {
      version: '1.0',
      encoding: 'UTF-8',
      standalone: true,
    });
    sst.att(
      'xmlns',
      'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    );
    sst.att('count', '' + this.arr.length);
    sst.att('uniqueCount', '' + this.arr.length);
    for (let i = 0; i <= this.arr.length; i++) {
      const si = sst.ele('si');
      si.ele('t', this.arr[i]);
      si.ele('phoneticPr', { fontId: 1, type: 'noConversion' });
    }
    return sst.end({ pretty: false });
  }
}
