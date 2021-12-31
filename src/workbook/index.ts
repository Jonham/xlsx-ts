import { writeFile } from 'fs'
import JSZip from 'jszip'
import { baseXl } from '../const/baseXl'
import { _Image, _Media } from '../types'
import { XlDrawingRels } from '../util/XlDrawingRels'
import { XlWorksheetRels } from '../util/XlWorksheetRels'
import { Sheet } from './Sheet'

export class Workbook {
  id: string
  filePath: string
  fileName: string

  sheets: Sheet[] = []
  medias: _Media[] = []

  // TODO Type
  /** @ss = new SharedStrings */
  sharedStrings: any
  //

  // TODO Type
  /** @cc: new ContentTypes(@) */
  contentType: any

  // TODO Type
  /** @da = new DocPropsApp(@) */
  docPropsApp: any

  // TODO Type
  /** @wb = new XlWorkbook(@) */
  XlWorkbook: any

  // TODO Type
  /** @wbre = new XlWorkbookRels(@) */
  XlWorkbookRels: any

  // TODO Type
  /** @st = new Style(@) */
  style: any

  // TODO Type
  /** @cc = new CalcChain(@) */
  calcChain: any

  // TODO Type
  /** @wsre = new XlSheetRels(@) */
  XlSheetRels: any

  // TODO Type
  /** @dw = new XlDrawing(@) */
  XlDrawing: any

  // TODO Type
  /** @dwre = new XlDrawingRels(@) */
  XlDrawingRels: any

  constructor(filePath: string, fileName: string) {
    this.filePath = filePath
    this.fileName = fileName
    this.id = (Math.random() * 9999999).toFixed(0)
    // # create temp folder & copy template data
    // # init
  }

  createSheet(name: string, cols: number, rows: number) {
    const sheet = new Sheet(this, name, cols, rows)
    this.sheets.push(sheet)
    return sheet
  }

  _addMediaFromImage(image: _Image) {
    // converts image into proper media data structure
    this.medias.push({ image })
  }

  _removeMediaFromImage(image: _Image) {
    // find image.
    const foundIndex = this.medias.findIndex(
      media => media.image.id === image.id
    )
    // remove it if found.
    if (foundIndex !== -1) {
      this.medias.splice(foundIndex, 1)
    }
  }

  // declare
  save(cb: Function): void
  save(target: string, cb: Function): void
  save(target: string, opts: any, cb: Function): void
  save(target: string | Function, opts: any = {}, cb?: Function) {
    if (typeof target === 'function' && !opts && !cb) {
      cb = target
      target = `${this.filePath}/${this.fileName}`
      opts = {}
    }
    if (typeof opts === 'function' && !cb) {
      cb = opts
      opts = {}
    }
    this._save(target as string, opts, cb!)
  }

  // private for save()
  private _save(target: string, opts: any = {}, cb: Function) {
    this.generate((err: Error, zip: any) => {
      let buffer
      let args = { type: 'nodebuffer' } as {
        type: 'nodebuffer'
        compressed?: 'DEFLATE'
      }
      if (opts.compressed) {
        args.compressed = 'DEFLATE'
      }

      buffer = zip.generateAsync(args).then((buffer: any) => {
        if (err) return cb(err)
        writeFile(target, buffer, err => cb(err))
      })
    })
  }

  // takes a callback function(err, zip) and returns a JSZip object on success
  generate(cb: Function) {
    const zip = new JSZip()

    for (const [key, value] of Object.entries(baseXl)) {
      zip.file(key, value)
    }

    // # 1 - build [Content_Types].xml
    zip.file('[Content_Types].xml', this.contentType.toxml())
    // # 2 - build docProps/app.xml
    zip.file('docProps/app.xml', this.docPropsApp.toxml())
    // # 3 - build xl/workbook.xml
    zip.file('xl/workbook.xml', this.XlWorkbook.toxml())
    // # 4 - build xl/sharedStrings.xml
    zip.file('xl/sharedStrings.xml', this.sharedStrings.toxml())
    // # 5 - build xl/_rels/workbook.xml.rels
    zip.file('xl/_rels/workbook.xml.rels', this.XlWorkbookRels.toxml())

    let wbMediaCounter = 1 // workbook media counter, per generation

    for (let i = 0; i < this.sheets.length; i++) {
      const sheet = this.sheets[i]

      sheet.wsRels = []

      for (let j = 0; j < sheet.images.length; j++) {
        const image = sheet.images[j]

        const dwRels = [] // drawing media list, for this release, one per drawing
        const relId = 'rId' + (sheet.wsRels.length + 1) // same for both media to drawing OR sheet to drawing rels.

        // - build xl/media/image(1-N).xml
        const mediaFilename = [wbMediaCounter, '.', image.extension].join('')
        zip.file(`xl/media/image${mediaFilename}`, image.content, {
          base64: true,
        })
        dwRels.push({ id: relId, target: `../media/image${mediaFilename}` })

        // - build xl/drawings/drawing(1-N).xml
        const drawingFilename = `${wbMediaCounter}.xml`
        sheet.wsRels.push({
          id: relId,
          target: `../drawings/drawing${drawingFilename}`,
        })
        zip.file(
          `xl/drawings/drawing${drawingFilename}`,
          image.toDrawingXml(relId, image)
        )

        // - build xl/drawings/_rels/drawing(1-N).xml.rels
        // TODO param type
        zip.file(
          `xl/drawings/_rels/drawing${wbMediaCounter}.xml.rels`,
          new XlDrawingRels(dwRels).toxml() as any
        )

        wbMediaCounter++
      }

      // - build xl/worksheets/_rels/sheet(1-N).xml.rels
      // TODO param type
      zip.file(
        `xl/worksheets/_rels/sheet${i + 1}.xml.rels`,
        new XlWorksheetRels(sheet.wsRels).toxml()
      )

      // - build xl/worksheets/sheet(1-N).xml
      // TODO param type
      zip.file(`xl/worksheets/sheet${i + 1}.xml`, this.sheets[i].toxml())
    }

    // 7 - build xl/styles.xml
    // TODO param type
    zip.file('xl/styles.xml', this.style.toxml())

    // 8 - build xl/calcChain.xml
    if (Object.keys(this.calcChain.cache).length > 0) {
      zip.file('xl/calcChain.xml', this.calcChain.toxml())
    }
    // 9 - build xl/worksheets/sheet(1-N).xml

    cb(null, zip)
  }

  cancel() {
    // delete temp folder
    console.error('workbook.cancel() is deprecated')
  }
}
