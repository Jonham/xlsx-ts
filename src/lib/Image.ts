import xmlbuilder from 'xmlbuilder'
import { TODO } from '../types'
import { Sheet } from '../workbook/Sheet'

type _options = TODO & { stretch: string }
type _rangeModel = {
  nativeCol: number
  nativeColOff: number
  nativeRow: number
  nativeRowOff: number
}
type _range = TODO & {
  from: { model: _rangeModel }
  to: { model: _rangeModel }
}
class Image {
  // TODO types
  id: TODO
  extension: TODO
  content: TODO
  range: _range
  options: _options

  editAs: string

  constructor(
    id: TODO,
    extension: TODO,
    content: TODO,
    range: _range,
    options: _options
  ) {
    this.id = id
    this.extension = extension
    this.content = content
    this.range = range
    this.options = options

    this.editAs = 'oneCell'
  }

  // TODO
  /**
   *
   * Inject image, it's used by sheet
   * 1. write data to media folder
   *   convert base 64 to text
   *   define filename based on number of existing medias
   *   writes to media
   * 2. create reference for media to drawing
   * 3. create the actual drawing using reference for media and set location
   * 4. creates reference for drawing to sheet.
   * 5. use image rel to sheet.
   */
  publish(sheet: Sheet, zip: TODO, media: TODO) {}

  toDrawingXml(relId: string, spec: TODO): string {
    // pngVersionRel = 'rId1'
    // svgVersionRel = 'rId2'

    const dr = xmlbuilder.create('xdr:wsDr', {
      version: '1.0',
      encoding: 'UTF-8',
      standalone: true,
    })
    // dr.att('xmlns:xdr', 'http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing')
    dr.att(
      'xmlns:xdr',
      'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
    )
    // dr.att('xmlns:a', 'http://purl.oclc.org/ooxml/drawingml/main')
    dr.att('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main')

    const twoCellAnchor = dr.ele('xdr:twoCellAnchor', { editAs: this.editAs })

    const _from = twoCellAnchor.ele('xdr:from') // {col: 0, colOff: 0, row: 1, rowOff: 0})
    _from.ele('xdr:col', this.range.from.model.nativeCol)
    _from.ele('xdr:colOff', this.range.from.model.nativeColOff)
    _from.ele('xdr:row', this.range.from.model.nativeRow)
    _from.ele('xdr:rowOff', this.range.from.model.nativeRowOff)

    const _to = twoCellAnchor.ele('xdr:to') // {col: 5, colOff: 419100, row: 16, rowOff: 0})
    _to.ele('xdr:col', this.range.to.model.nativeCol)
    _to.ele('xdr:colOff', this.range.to.model.nativeColOff)
    _to.ele('xdr:row', this.range.to.model.nativeRow)
    _to.ele('xdr:rowOff', this.range.to.model.nativeRowOff)

    const pic = twoCellAnchor.ele('xdr:pic')
    const nvPicPr = pic.ele('xdr:nvPicPr')
    // Graphic name printed on xml are not seen or used as reference
    // Skiping proper calculation part and just finding a random until 100000
    const graphic_index = 30924 || Math.round(Math.random() * 100000)
    const cNvPr = nvPicPr.ele('xdr:cNvPr', {
      id: 3,
      name: 'Graphic ' + graphic_index,
    })
    cNvPr
      .ele('a:extLst')
      .ele('a:ext', { uri: '{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}' })
      .ele('a16:creationId', {
        'xmlns:a16': 'http://schemas.microsoft.com/office/drawing/2014/main',
        id: '{9D66B5F7-2553-484C-A5BE-4D0B8D57E08B}',
      })
    nvPicPr.ele('xdr:cNvPicPr').ele('a:picLocks', { noChangeAspect: 1 })

    const blipFill = pic.ele('xdr:blipFill')

    const blipModel = {
      'xmlns:r':
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'r:embed': '',
      cstate: '',
    }

    if (relId) {
      blipModel['r:embed'] = relId
    }

    blipModel['cstate'] = 'print'
    const blip = blipFill.ele('a:blip', blipModel)
    const extLst = blip.ele('a:extLst')
    const ext = extLst.ele('a:ext', {
      uri: '{28A0092B-C50C-407E-A947-70E740481C1C}',
    })
    ext.ele('a14:useLocalDpi', {
      'xmlns:a14': 'http://schemas.microsoft.com/office/drawing/2010/main',
      val: 0,
    })

    if (this.extension == 'svg') {
      const ext = extLst.ele('a:ext', {
        uri: '{96DAC541-7B7A-43D3-8B79-37D633B846F1}',
      })
      ext.ele('asvg:svgBlip', {
        'xmlns:asvg':
          'http://schemas.microsoft.com/office/drawing/2016/SVG/main',
        'r:embed': relId,
      })
    }

    if (this.options.stretch) {
      blipFill.ele('a:stretch')
    }
    blipFill.ele('srcRect')
    //.ele('a:fillRect')

    const spPr = pic.ele('xdr:spPr')
    const xfrm = spPr.ele('a:xfrm')
    xfrm.ele('a:off', { x: 609600, y: 190500 })
    xfrm.ele('a:ext', { cx: 2857500, cy: 2857500 })
    spPr.ele('a:prstGeom', { prst: 'rect' }).ele('a:avLst')
    twoCellAnchor.ele('xdr:clientData')
    return dr.end({ pretty: false })
  }
}
