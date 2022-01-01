import xml from 'xmlbuilder';
import { _dwRel } from '../types';

export class XlDrawingRels {
  dwRels: _dwRel[];

  constructor(dwRels: _dwRel[]) {
    this.dwRels = dwRels;
  }
  generate() {}
  toxml(): string {
    const rs = xml.create('Relationships', {
      version: '1.0',
      encoding: 'UTF-8',
      standalone: true,
    });
    rs.att(
      'xmlns',
      'http://schemas.openxmlformats.org/package/2006/relationships',
    );
    for (const dwRel of this.dwRels) {
      rs.ele('Relationship', {
        Id: dwRel.id,
        // Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        Target: dwRel.target,
      });
    }

    return rs.end({ pretty: false });
  }
}
