import xmlbuilder from 'xmlbuilder';
import { _wsRel } from '../types';

export class XlWorksheetRels {
  wsRels: _wsRel[] = [];
  constructor(wsRels: _wsRel[]) {
    this.wsRels = wsRels;
  }
  generate() {}

  toxml() {
    const rs = xmlbuilder.create('Relationships', {
      version: '1.0',
      encoding: 'UTF-8',
      standalone: true,
    });
    rs.att(
      'xmlns',
      'http://schemas.openxmlformats.org/package/2006/relationships',
    );
    for (const wsRel of this.wsRels) {
      rs.ele('Relationship', {
        Id: wsRel.id,
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
        Target: wsRel.target,
      });
    }
    return rs.end({ pretty: false });
  }
}
