import { TODO, _Fill, _Font, _FONT_ID, _Border } from "../../types"
import { numberFormats } from './numberFormats'

class Style {
  numberFormats: Record<number, string> = {...numberFormats}

  // TODO
  book: TODO
  // font.toString cache
  cache: Record<string, _FONT_ID>
  mfonts: _Font[]
  mfills: _Fill[]
  mbders: _Border[]
  mstyle: TODO
  numFmtNextId: number

  constructor (book: TODO) {
    this.book = book

    this.cache = {}
    this.mfonts = [] // font style
    this.mfills = [] // fill style
    this.mbders = [] // border style
    this.mstyle = [] // cell style<ref-font,ref-fill,ref-border,align>
    this.numFmtNextId = 164

    this.with_default()
  }

  def_font_id: TODO
  def_fill_id: TODO
  def_bder_id: TODO
  def_align: TODO
  def_valign: TODO
  def_rotate: TODO
  def_wrap: TODO
  def_numfmt_id: TODO
  def_style_id: TODO
  with_default () {
    this.def_font_id = this.font2id()
    this.def_fill_id = this.fill2id()
    this.def_bder_id = this.bder2id()
    this.def_align = '-'
    this.def_valign = '-'
    this.def_rotate = '-'
    this.def_wrap = '-'
    this.def_numfmt_id = 0
    this.def_style_id = this.style2id({
      font_id: this.def_font_id,
      fill_id: this.def_fill_id,
      bder_id: this.def_bder_id,
      align: this.def_align,
      valign: this.def_valign,
      rotate: this.def_rotate
    })
  }

  font2id (font: Partial<_Font> = {}) {
    // Default
    const defaultFont: Partial<_Font> = {
      bold: '-',
      iter: '-',
      sz: '11',
      color: '-',
      name: 'Calibri',
      scheme: 'minor',
      family: '2',
      underline: '-',
      strike: '-',
      outline: '-',
      shadow: '-',
    }
    font = {
      ...defaultFont,
      ...(font as any || {}),
    }
    
    const str = 'font_' + font.bold + font.iter + font.sz + font.color + font.name + font.scheme + font.family + font.underline + font.strike + font.outline + font.shadow
    const id = this.cache[str]
    if (id) {return id}
    else {
      this.mfonts.push(font as _Font)
      this.cache[str] = this.mfonts.length
      return this.mfonts.length
    }
  }

  fill2id (fill: Partial<_Fill> = {}) {
    const defaultFill: Partial<_Fill> = {
      type: 'none',
      bgColor: '-',
      fgColor: '-',
    }
    fill = {
      ...defaultFill,
      ...fill
    }

    const str = 'fill_' + fill.type + fill.bgColor + fill.fgColor
    const id = this.cache[str]
    if (id) {
      return id
    }
    else {
      this.mfills.push(fill as _Fill)
      this.cache[str] = this.mfills.length
      return this.mfills.length
    }
  }

  bder2id (border: Partial<_Border> = {}) {
    const defaultBorder: _Border = {
      left: '-',
      right: '-',
      top: '-',
      bottom: '-',
    }

    border = {
      ...defaultBorder,
      ...border
    }

    const {
      left, right, top, bottom,
    } = border

    const str = JSON.stringify(["bder_",left, right, top, bottom])
    const id = this.cache[str]
    if (id) {
      return id
    }
    else {
      this.mbders.push(border as _Border)
      this.cache[str] = this.mbders.length
      return this.mbders.length
    }
  }

  numfmt2id (numfmt: number | string) {
    if (typeof numfmt == 'number') {
      return numfmt
    } else if (typeof numfmt == 'string') {
      if (!numfmt) {
        throw "Invalid format specification"
      }

      for (const [key, value] of Object.entries(this.numberFormats)) {
        if (value === numfmt) {
          return parseInt(key)
        }
      }

      // if it's not in numberFormats, we parse the string and add it the end of numberFormats
      // numfmt = numfmt
      //   .replace(/&/g, '&amp')
      //   .replace(/</g, '&lt;')
      //   .replace(/>/g, '&gt;')
      //   .replace(/"/g, '&quot;')
      this.numberFormats[++this.numFmtNextId] = numfmt
      return this.numFmtNextId
    }
  }

  style2id (style = {}) {

  }
  // style2id: (style)->
  //   style.align or= @def_align
  //   style.valign or= @def_valign
  //   style.rotate or= @def_rotate
  //   style.wrap or= @def_wrap
  //   style.font_id or= @def_font_id
  //   style.fill_id or= @def_fill_id
  //   style.bder_id or= @def_bder_id
  //   style.numfmt_id or= @def_numfmt_id
  //   k = 's_' + [style.font_id, style.fill_id, style.bder_id, style.align, style.valign, style.wrap, style.rotate,
  //     style.numfmt_id].join('_')
  //   id = @cache[k]
  //   if id
  //     return id
  //   else
  //     @mstyle.push style
  //     @cache[k] = @mstyle.length
  //     return @mstyle.length

  // toxml: ()->
  //   ss = xml.create('styleSheet', {version: '1.0', encoding: 'UTF-8', standalone: true})
  //   ss.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
  //   # add all numFmts >= 164 as <numFmt numFmtId="${o.num_fmt_id}" formatCode="numFmt"/>
  //   customNumFmts = [];
  //   for key, fmt of @numberFormats
  //     if parseInt(key) >= 164
  //       customNumFmts.push({numFmtId: key, formatCode: fmt});
  //   if customNumFmts.length > 0
  //     numFmts = ss.ele('numFmts', {
  //       count: customNumFmts.length
  //     });
  //     for o in customNumFmts
  //       numFmts.ele('numFmt', o)
  //   fonts = ss.ele('fonts', {count: @mfonts.length})
  //   for o in @mfonts
  //     e = fonts.ele('font')
  //     e.ele('b') if o.bold isnt '-'
  //     e.ele('i') if o.iter isnt '-'
  //     e.ele('u') if o.underline isnt '-'
  //     e.ele('strike') if o.strike isnt '-'
  //     e.ele('outline') if o.outline isnt '-'
  //     e.ele('shadow') if o.shadow isnt '-'

  //     e.ele('sz', {val: o.sz})
  //     e.ele('color', {rgb: o.color}) if o.color isnt '-'
  //     e.ele('name', {val: o.name})
  //     e.ele('family', {val: o.family})
  //     e.ele('charset', {val: '134'})
  //     e.ele('scheme', {val: 'minor'}) if o.scheme isnt '-'
  //   fills = ss.ele('fills', {count: @mfills.length + 2})
  //   fills.ele('fill').ele('patternFill', {patternType: 'none'})
  //   fills.ele('fill').ele('patternFill', {patternType: 'gray125'})
  //   #<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>

  //   for o in @mfills
  //     e = fills.ele('fill')
  //     es = e.ele('patternFill', {patternType: o.type})
  //     es.ele('fgColor', {rgb: o.fgColor}) if o.fgColor isnt '-'
  //     es.ele('bgColor', {indexed: o.bgColor}) if o.bgColor isnt '-'
  //   bders = ss.ele('borders', {count: @mbders.length})
  //   for o in @mbders

  //     e = bders.ele('border')

  //     if o.left isnt '-'
  //       if typeof o.left is 'string'
  //         e.ele('left', {style: o.left}).ele('color', {auto: '1'})
  //       else
  //         e.ele('left', {style: o.left.style}).ele('color', o.left.color)
  //     else e.ele('left')

  //     if o.right isnt '-'
  //       if typeof o.right is 'string'
  //         e.ele('right', {style: o.right}).ele('color', {auto: '1'})
  //       else
  //         e.ele('right', {style: o.right.style}).ele('color', o.right.color)
  //     else e.ele('right')

  //     if o.top isnt '-'
  //       if typeof o.top is 'string'
  //         e.ele('top', {style: o.top}).ele('color', {auto: '1'})
  //       else
  //         e.ele('top', {style: o.top.style}).ele('color', o.top.color)
  //     else e.ele('top')

  //     if o.bottom isnt '-'
  //       if typeof o.bottom is 'string'
  //         e.ele('bottom', {style: o.bottom}).ele('color', {auto: '1'})
  //       else
  //         e.ele('bottom', {style: o.bottom.style}).ele('color', o.bottom.color)
  //     else e.ele('bottom')



  //     e.ele('diagonal')
  //   ss.ele('cellStyleXfs', {count: '1'}).ele('xf', {
  //     numFmtId: '0',
  //     fontId: '0',
  //     fillId: '0',
  //     borderId: '0'
  //   }).ele('alignment', {vertical: 'center'})
  //   cs = ss.ele('cellXfs', {count: @mstyle.length})
  //   for o in @mstyle
  //     e = cs.ele('xf', {
  //       numFmtId: o.numfmt_id || '0',
  //       fontId: (o.font_id - 1),
  //       fillId: o.fill_id + 1,
  //       borderId: (o.bder_id - 1),
  //       xfId: '0'
  //     })
  //     e.att('applyFont', '1') if o.font_id isnt 1
  //     e.att('applyFill', '1') if o.fill_id isnt 1
  //     e.att('applyNumberFormat', '1') if o.numfmt_id isnt undefined
  //     e.att('applyBorder', '1') if o.bder_id isnt 1
  //     if o.align isnt '-' or o.valign isnt '-' or o.wrap isnt '-'
  //       e.att('applyAlignment', '1')
  //       ea = e.ele('alignment', {
  //         textRotation: (if o.rotate is '-' then '0' else o.rotate),
  //         horizontal: (if o.align is '-' then 'left' else o.align),
  //         vertical: (if o.valign is '-' then 'bottom' else o.valign)
  //       })
  //       ea.att('wrapText', '1') if o.wrap isnt '-'
  //   ss.ele('cellStyles', {count: '1'}).ele('cellStyle', {name: 'Normal', xfId: '0', builtinId: '0'})
  //   ss.ele('dxfs', {count: '0'})
  //   ss.ele('tableStyles', {count: '0', defaultTableStyle: 'TableStyleMedium9', defaultPivotStyle: 'PivotStyleLight16'})
  //   return ss.end({pretty: false})
