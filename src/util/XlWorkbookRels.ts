import xmlbuilder from "xmlbuilder"
import { Workbook } from "../workbook"

export class XlWorkbookRels {
  book: Workbook
  constructor (book: Workbook) {
    this.book = book
  }


  toxml () {
    const rs = xmlbuilder.create('Relationships', {version: '1.0', encoding: 'UTF-8', standalone: true})
    rs.att('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships')
    for (let i = 1; i <= this.book.sheets.length ;i++) {
      rs.ele('Relationship', {
        Id: 'rId' + i,
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
        Target: 'worksheets/sheet' + i + '.xml'
      })
    }
    rs.ele('Relationship', {
      Id: 'rId' + (this.book.sheets.length + 1),
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
      Target: 'theme/theme1.xml'
    })
    rs.ele('Relationship', {
      Id: 'rId' + (this.book.sheets.length + 2),
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
      Target: 'styles.xml'
    })
    rs.ele('Relationship', {
      Id: 'rId' + (this.book.sheets.length + 3),
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
      Target: 'sharedStrings.xml'
    })
    rs.ele('Relationship', {
      Id: 'rId' + (this.book.sheets.length + 4),
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain',
      Target: 'calcChain.xml'
    });
    return rs.end()
    // if (Object.getOwnPropertyNames(this.book.cc.cache).length > 0) {
    //   rs.ele('Relationship', {
    //     Id: 'rId' + (this.book.sheets.length + 4),
    //     Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain',
    //     Target: 'calcChain.xml'
    //   });
    // }
    // return rs.end({pretty: false})
  }

}
