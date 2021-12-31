###
  MS Excel 2007 Creameater v0.0.1
  Author : chuanyi.zheng@gmail.com
  Extended: pieter@protobi.com
  History: 2012/11/07 first created
###

if (window? && window.JSZip?)
  JSZip = window.JSZip
else if (typeof require != 'undefined')
  JSZip = require 'jszip'
else
  throw ("JSZip not defined")

if (window? && window.xmlbuilder?)
  xml = window.xmlbuilder
else if (typeof require != 'undefined')
  xml = require 'xmlbuilder'
else
  throw ("xmlbuilder not defined")

if (window? && window.xmlbuilder?)
  fs = window.xmlbuilder
else if (typeof require != 'undefined')
  fs = require 'fs'

####tool =
#  i2a : (i) ->
#    return 'ABCDEFGHIJKLMNOPQ###RSTUVWXYZ123'.charAt(i-1)

tool =
  i2a: (column) ->
    temp = undefined
    letter = ''
    while column > 0
      temp = (column - 1) % 26
      letter = String.fromCharCode(temp + 65) + letter
      column = (column - temp - 1) / 26
    return letter

ImageType =
  SVG: "image/svg+xml"
  PNG: "image/png"
# JPEG : 401

Function::property = (prop, desc) ->
  Object.defineProperty @prototype, prop, desc

addressRegex = /^[A-Z]+\d+$/
# =========================================================================
# Column Letter to Number conversion
colCache =
  _dictionary: [
    'A'
    'B'
    'C'
    'D'
    'E'
    'F'
    'G'
    'H'
    'I'
    'J'
    'K'
    'L'
    'M'
    'N'
    'O'
    'P'
    'Q'
    'R'
    'S'
    'T'
    'U'
    'V'
    'W'
    'X'
    'Y'
    'Z'
  ]
  _l2nFill: 0
  _l2n: {}
  _n2l: []
  _level: (n) ->
    if n <= 26
      return 1
    if n <= 26 * 26
      return 2
    3
  _fill: (level) ->
    c = undefined
    v = undefined
    l1 = undefined
    l2 = undefined
    l3 = undefined
    n = 1
    if level >= 4
      throw new Error('Out of bounds. Excel supports columns from 1 to 16384')
    if @_l2nFill < 1 and level >= 1
      while n <= 26
        c = @_dictionary[n - 1]
        @_n2l[n] = c
        @_l2n[c] = n
        n++
      @_l2nFill = 1
    if @_l2nFill < 2 and level >= 2
      n = 27
      while n <= 26 + 26 * 26
        v = n - (26 + 1)
        l1 = v % 26
        l2 = Math.floor(v / 26)
        c = @_dictionary[l2] + @_dictionary[l1]
        @_n2l[n] = c
        @_l2n[c] = n
        n++
      @_l2nFill = 2
    if @_l2nFill < 3 and level >= 3
      n = 26 + 26 * 26 + 1
      while n <= 16384
        v = n - (26 * 26 + 26 + 1)
        l1 = v % 26
        l2 = Math.floor(v / 26) % 26
        l3 = Math.floor(v / (26 * 26))
        c = @_dictionary[l3] + @_dictionary[l2] + @_dictionary[l1]
        @_n2l[n] = c
        @_l2n[c] = n
        n++
      @_l2nFill = 3
    return
  l2n: (l) ->
    if !@_l2n[l]
      @_fill l.length
    if !@_l2n[l]
      throw new Error('Out of bounds. Invalid column letter: ' + l)
    @_l2n[l]
  n2l: (n) ->
    if n < 1 or n > 16384
      throw new Error(n + ' is out of bounds. Excel supports columns from 1 to 16384')
    if !@_n2l[n]
      @_fill @_level(n)
    @_n2l[n]
  _hash: {}
  validateAddress: (value) ->
    if !addressRegex.test(value)
      throw new Error('Invalid Address: ' + value)
    true
  decodeAddress: (value) ->
    addr = value.length < 5 and @_hash[value]
    if addr
      return addr
    hasCol = false
    col = ''
    colNumber = 0
    hasRow = false
    row = ''
    rowNumber = 0
    i = 0
    char = undefined
    while i < value.length
      char = value.charCodeAt(i)
      # col should before row
      if !hasRow and char >= 65 and char <= 90
# 65 = 'A'.charCodeAt(0)
# 90 = 'Z'.charCodeAt(0)
        hasCol = true
        col += value[i]
        # colNumber starts from 1
        colNumber = colNumber * 26 + char - 64
      else if char >= 48 and char <= 57
# 48 = '0'.charCodeAt(0)
# 57 = '9'.charCodeAt(0)
        hasRow = true
        row += value[i]
        # rowNumber starts from 0
        rowNumber = rowNumber * 10 + char - 48
      else if hasRow and hasCol and char != 36
# 36 = '$'.charCodeAt(0)
        break
      i++
    if !hasCol
      colNumber = undefined
    else if colNumber > 16384
      throw new Error('Out of bounds. Invalid column letter: ' + col)
    if !hasRow
      rowNumber = undefined
    # in case $row$col
    value = col + row
    address =
      address: value
      col: colNumber
      row: rowNumber
      $col$row: '$' + col + '$' + row
    # mem fix - cache only the tl 100x100 square
    if colNumber <= 100 and rowNumber <= 100
      @_hash[value] = address
      @_hash[address.$col$row] = address
    address
  getAddress: (r, c) ->
    if c
      address = @n2l(c) + r
      return @decodeAddress(address)
    @decodeAddress r
  decode: (value) ->
    parts = value.split(':')
    if parts.length == 2
      tl = @decodeAddress(parts[0])
      br = @decodeAddress(parts[1])
      result =
        top: Math.min(tl.row, br.row)
        left: Math.min(tl.col, br.col)
        bottom: Math.max(tl.row, br.row)
        right: Math.max(tl.col, br.col)
      # reconstruct tl, br and dimensions
      result.tl = @n2l(result.left) + result.top
      result.br = @n2l(result.right) + result.bottom
      result.dimensions = result.tl + ':' + result.br
      return result
    @decodeAddress value
  decodeEx: (value) ->
    groups = value.match(/(?:(?:(?:'((?:[^']|'')*)')|([^'^ !]*))!)?(.*)/)
    sheetName = groups[1] or groups[2]
    # Qouted and unqouted groups
    reference = groups[3]
    # Remaining address
    parts = reference.split(':')
    if parts.length > 1
      tl = @decodeAddress(parts[0])
      br = @decodeAddress(parts[1])
      top = Math.min(tl.row, br.row)
      left = Math.min(tl.col, br.col)
      bottom = Math.max(tl.row, br.row)
      right = Math.max(tl.col, br.col)
      tl = @n2l(left) + top
      br = @n2l(right) + bottom
      return {
        top: top
        left: left
        bottom: bottom
        right: right
        sheetName: sheetName
        tl:
          address: tl
          col: left
          row: top
          $col$row: '$' + @n2l(left) + '$' + top
          sheetName: sheetName
        br:
          address: br
          col: right
          row: bottom
          $col$row: '$' + @n2l(right) + '$' + bottom
          sheetName: sheetName
        dimensions: tl + ':' + br
      }
    if reference.startsWith('#')
      return if sheetName then {
        sheetName: sheetName,
        error: reference
      } else {error: reference}
    address = @decodeAddress(reference)
    if sheetName then {
      sheetName: sheetName,
      address: address.address,
      col: address.col,
      row: address.row,
      $col$row: '$' + col + '$' + row
    } else address
  encodeAddress: (row, col) ->
    colCache.n2l(col) + row
  encode: ->
    switch arguments.length
      when 2
        return colCache.encodeAddress(arguments[0], arguments[1])
      when 4
        return colCache.encodeAddress(arguments[0], arguments[1]) + ':' + colCache.encodeAddress(arguments[2], arguments[3])
      else
        throw new Error('Can only encode with 2 or 4 arguments')
    return
  inRange: (range, address) ->
    left = range[0]
    top = range[1]
    right = range[range.length - 2]
    bottom = range[range.length - 1]
    # const [left, top, , right, bottom] = range;
    col = address[0]
    row = address[1]
    col >= left and col <= right and row >= top and row <= bottom
# ---


class ContentTypes
  constructor: (@book)->
  _getKnowImageTypes: () ->
    imagesToAdd = []
    debugger;
    imagesToAddDistinct = {}
    for sheet in @book.sheets
      for image in sheet.images
        unless not imagesToAddDistinct[image.extension]
          imagesToAdd.push({Extension: image.extension, ContentType: image.contentType})
    imagesToAdd

  toxml: ()->
    types = xml.create('Types', {version: '1.0', encoding: 'UTF-8', standalone: true})
    types.att('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types')


    # TODO remove fixed types here
    types.ele('Default', {Extension: 'png', ContentType: 'image/png'})
    types.ele('Default', {Extension: 'svg', ContentType: 'image/svg+xml'})
    types.ele('Default', {Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml'})
    types.ele('Default', {Extension: 'xml', ContentType: 'application/xml'})

    types.ele('Override', {
      PartName: '/xl/workbook.xml',
      ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
    })
    for i in [1..@book.sheets.length]
      types.ele('Override', {
        PartName: '/xl/worksheets/sheet' + i + '.xml',
        ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
      })

    types.ele('Override', {
      PartName: '/xl/theme/theme1.xml',
      ContentType: 'application/vnd.openxmlformats-officedocument.theme+xml'
    })
    types.ele('Override', {
      PartName: '/xl/calcChain.xml',
      ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml'
    });
    types.ele('Override', {
      PartName: '/xl/styles.xml',
      ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
    })
    for sheet in @book.sheets
      for image in sheet.images
        types.ele('Override', {
          PartName: '/xl/drawings/drawing' + image.id + '.xml',
          ContentType: 'application/vnd.openxmlformats-officedocument.drawing+xml'
        })
    types.ele('Override', {
      PartName: '/docProps/core.xml',
      ContentType: 'application/vnd.openxmlformats-package.core-properties+xml'
    })
    types.ele('Override', {
      PartName: '/docProps/app.xml',
      ContentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
    })

#    if Object.getOwnPropertyNames(@book.cc.cache).length > 0
#      types.ele('Override', {
#        PartName: '/xl/calcChain.xml',
#        ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml'
#      });

    # for knowImageType in @_getKnowImageTypes
    # 	types.ele('Default', knowImageType)
    types.ele('Override', {
      PartName: '/xl/sharedStrings.xml',
      ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
    })

    return types.end({pretty: false})

class DocPropsApp
  constructor: (@book)->

  toxml: ()->
    props = xml.create('Properties', {version: '1.0', encoding: 'UTF-8', standalone: true})
    props.att('xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties')
    props.att('xmlns:vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes')
    props.ele('Application', 'Microsoft Excel')
    props.ele('DocSecurity', '0')
    props.ele('ScaleCrop', 'false')
    tmp = props.ele('HeadingPairs').ele('vt:vector', {size: 2, baseType: 'variant'})
    tmp.ele('vt:variant').ele('vt:lpstr', 'Worksheets')
    tmp.ele('vt:variant').ele('vt:i4', '' + @book.sheets.length)
    tmp = props.ele('TitlesOfParts').ele('vt:vector', {size: @book.sheets.length, baseType: 'lpstr'})
    for i in [1..@book.sheets.length]
      tmp.ele('vt:lpstr', @book.sheets[i - 1].name)
    props.ele('Company')
    props.ele('LinksUpToDate', 'false')
    props.ele('SharedDoc', 'false')
    props.ele('HyperlinksChanged', 'false')
    props.ele('AppVersion', '12.0000')
    return props.end({pretty: false})

class XlWorkbook
  constructor: (@book)->

  toxml: ()->

    wb = xml.create('workbook', {version: '1.0', encoding: 'UTF-8', standalone: true})
    wb.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    wb.att('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    wb.att('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
    wb.att('mc:Ignorable', "x15 xr xr6 xr10 xr2")
    wb.att('xmlns:x15', "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main")
    wb.att('xmlns:xr', "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
    wb.att('xmlns:xr6', "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6")
    wb.att('xmlns:xr10', "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10")
    wb.att('xmlns:xr2', "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")

    wb.ele('fileVersion', {appName: 'xl', lastEdited: '4', lowestEdited: '4', rupBuild: '4505'})
    wb.ele('workbookPr', {filterPrivacy: '1', defaultThemeVersion: '124226'})
    wb.ele('bookViews').ele('workbookView', {xWindow: '0', yWindow: '90', windowWidth: '19200', windowHeight: '11640'})

    tmp = wb.ele('sheets')
    for i in [1..@book.sheets.length]
      tmp.ele('sheet', {name: @book.sheets[i - 1].name, sheetId: '' + i, 'r:id': 'rId' + i})


    definedNames = wb.ele('definedNames') # one entry per autofilter


    @book.sheets.forEach((sheet, idx) ->
      if (sheet.autofilter)
        definedNames.ele('definedName', {
          name: '_xlnm._FilterDatabase',
          hidden: '1',
          localSheetId: idx
        }).raw("'" + sheet.name + "'!" + sheet.getRange())

      if (sheet._repeatRows || sheet._repeatCols)
        range = ''
        if (sheet._repeatCols)
          range += "'" + sheet.name + "'!$" + tool.i2a(sheet._repeatCols.start) + ":$"+tool.i2a(sheet._repeatCols.end)
        if (sheet._repeatRows)
          range += ",'" + sheet.name + "'!$" + (sheet._repeatRows.start) + ":$"+(sheet._repeatRows.end)

        definedNames.ele('definedName', {
          name: "_xlnm.Print_Titles"
          localSheetId: idx
        }).raw(range)
    )



    wb.ele('calcPr', {calcId: '124519'})


    return wb.end({pretty: false})

class XlWorkbookRels
  constructor: (@book)->

  toxml: ()->
    rs = xml.create('Relationships', {version: '1.0', encoding: 'UTF-8', standalone: true})
    rs.att('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships')
    for i in [1..@book.sheets.length]
      rs.ele('Relationship', {
        Id: 'rId' + i,
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
        Target: 'worksheets/sheet' + i + '.xml'
      })
    rs.ele('Relationship', {
      Id: 'rId' + (@book.sheets.length + 1),
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
      Target: 'theme/theme1.xml'
    })
    rs.ele('Relationship', {
      Id: 'rId' + (@book.sheets.length + 2),
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
      Target: 'styles.xml'
    })
    rs.ele('Relationship', {
      Id: 'rId' + (@book.sheets.length + 3),
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
      Target: 'sharedStrings.xml'
    })
    rs.ele('Relationship', {
      Id: 'rId' + (this.book.sheets.length + 4),
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain',
      Target: 'calcChain.xml'
    });
    return rs.end()
    if Object.getOwnPropertyNames(@book.cc.cache).length > 0
      rs.ele('Relationship', {
        Id: 'rId' + (this.book.sheets.length + 4),
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain',
        Target: 'calcChain.xml'
      });
    return rs.end({pretty: false})

# class XlWorksheetRels
#   constructor: (@wsRels) ->
#   generate: () ->

#   toxml: () ->
#     rs = xml.create('Relationships', {version: '1.0', encoding: 'UTF-8', standalone: true})
#     rs.att('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships')
#     for wsRel in @wsRels
#       rs.ele('Relationship', {
#         Id: wsRel.id,
#         Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
#         Target: wsRel.target
#       })
#     rs.end({pretty: false})
# DONE
# class XlDrawingRels
#   constructor: (@dwRels) ->
#   generate: () ->

#   toxml: () ->
#     rs = xml.create('Relationships', {version: '1.0', encoding: 'UTF-8', standalone: true})
#     rs.att('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships')
#     for dwRel in @dwRels
#       rs.ele('Relationship', {
#         Id: dwRel.id,
# # Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
#         Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
#         Target: dwRel.target
#       })
#     rs.end({pretty: false})

class SharedStrings
  constructor: ()->
    @cache = {}
    @arr = []

  str2id: (s)->
    id = @cache[s]
    if id
      return id
    else
      @arr.push s
      @cache[s] = @arr.length
      return @arr.length

  toxml: ()->
    sst = xml.create('sst', {version: '1.0', encoding: 'UTF-8', standalone: true})
    sst.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    sst.att('count', '' + @arr.length)
    sst.att('uniqueCount', '' + @arr.length)
    for i in [0...@arr.length]
      si = sst.ele('si')
      si.ele('t', @arr[i])
      si.ele('phoneticPr', {fontId: 1, type: 'noConversion'})
    return sst.end({pretty: false})

class Anchor
  constructor: (@worksheet, address, offset) ->
    if offset == undefined
      offset = 0
    if !address
      @nativeCol = 0
      @nativeColOff = 0
      @nativeRow = 0
      @nativeRowOff = 0
    else if typeof address == 'string'
      decoded = colCache.decodeAddress(address)
      @nativeCol = decoded.col + offset
      @nativeColOff = 0
      @nativeRow = decoded.row + offset
      @nativeRowOff = 0
    else if address.nativeCol != undefined
      @nativeCol = address.nativeCol or 0
      @nativeColOff = address.nativeColOff || 0
      @nativeRow = address.nativeRow or 0
      @nativeRowOff = address.nativeRowOff || 0
    else if address.col != undefined
      @col = address.col + offset
      @row = address.row + offset
    else
      @nativeCol = 0
      @nativeColOff = 0
      @nativeRow = 0
      @nativeRowOff = 0
    return

  @property 'col', {
    get: -> @nativeCol + Math.min(@colWidth - 1, @nativeColOff) / @colWidth
    set: (v) ->
      @nativeCol = Math.floor(v)
      @nativeColOff = Math.floor((v - (@nativeCol)) * @colWidth)
      return
    enumerable: true
    configurable: true
  }
  @property 'row', {
    get: ->
      @nativeRow + Math.min(@rowHeight - 1, @nativeRowOff) / @rowHeight
    set: (v) ->
      @nativeRow = Math.floor(v)
      @nativeRowOff = Math.floor((v - (@nativeRow)) * @rowHeight)
      return
    enumerable: true
    configurable: true
  }
  @property 'colWidth', {
    get: ->
      if @worksheet and @worksheet.width(@nativeCol, @nativeCol + 1) then
#  and @worksheet.getColumn(@nativeCol + 1).isCustomWidth then
# Math.floor(@worksheet.getColumn(@nativeCol + 1).width * 10000) else 640000
    enumerable: true
    configurable: true
  }
  @property 'rowHeight', {
    get: ->
      if @worksheet and @worksheet.getRow(@nativeRow + 1) and @worksheet.getRow(@nativeRow + 1).height then Math.floor(@worksheet.getRow(@nativeRow + 1).height * 10000) else 180000
    enumerable: true
    configurable: true
  }
  @property 'model', {
    get: -> {
      nativeCol: @nativeCol
      nativeColOff: @nativeColOff
      nativeRow: @nativeRow
      nativeRowOff: @nativeRowOff
    },
    set: (value) ->
      @nativeCol = value.nativeCol
      @nativeColOff = value.nativeColOff
      @nativeRow = value.nativeRow
      @nativeRowOff = value.nativeRowOff
      return
    enumerable: true
    configurable: true
  }

  asInstance = (model) ->
    if model instanceof Anchor or model == null then model else new Anchor(model)

  toxml: (xml)->
    wb = xml.create('workbook', {version: '1.0', encoding: 'UTF-8', standalone: true})
    wb.ele('from')
      .ele('workbookView', {xWindow: '0', yWindow: '90', windowWidth: '19200', windowHeight: '11640'})

    wb.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    wb.att('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

    wb.ele('fileVersion', {appName: 'xl', lastEdited: '4', lowestEdited: '4', rupBuild: '4505'})
    wb.ele('workbookPr', {filterPrivacy: '1', defaultThemeVersion: '124226'})
    wb.ele('bookViews').ele('workbookView', {xWindow: '0', yWindow: '90', windowWidth: '19200', windowHeight: '11640'})



class Sheet
  constructor: (@book, @name, @cols, @rows) ->
    @name = @name.replace(/[/*:?\[\]]/g, '-')

    @data = {}
    for i in [1..@rows]
      @data[i] = {}
      for j in [1..@cols]
        @data[i][j] = {v: 0}
    @merges = []
    @col_wd = []
    @row_ht = {}
    @styles = {}
    @formulas=[]
    @_pageMargins= {left: '0.7', right: '0.7', top: '0.75', bottom: '0.75', header: '0.3', footer: '0.3'}
    @images = []

  # validates exclusivity between filling base64, filename, buffer properties.
  # validates extension is among supported types.
  # concurrency this is a critical path add semaphor, only one image can be added at the time.
  # there's a risk of adding image in parallel and returing diferent id between push and returning.
  # exceljs also contains same risk, despite of collecting id before.
  addImage: (image) ->
    if !image || !image.range || !image.base64 || !image.extension
      throw Error('please verify your image format')


    ## tries to decode range
    if (typeof image.range != 'string') || !/\w+\d+:\w+\d/i.test(image.range)
      throw Error('Please provide range parameter like `B2:F6`.')

    decoded = colCache.decode(image.range);
    @range = {
      from: new Anchor(@worksheet, decoded.tl, -1),
      to: new Anchor(@worksheet, decoded.br, 0),
      editAs: 'oneCell',
    }

    id = @book.medias.length + 1
    imageToAdd = new Image(id, image.extension, image.base64, @range, image.options||{})
    media = @book._addMediaFromImage(imageToAdd)
    # drawingId = @book._addDrawingFromImage(imageToAdd)
    # wsDwRelId = @sheet._addDrawingFromImage(imageToAdd)
    console.log(imageToAdd)
    @images.push(imageToAdd)

    return id

  getImage: (id) -> @images[id]

  getImages: () ->
    @images

  removeImage: (id) ->
    foundIndex = -1
    for image, index in @images
      if image.id == id
        foundIndex = index
        break
    if foundIndex > -1
      @images.splice(foundIndex, 1)

  ### old approach for adding background images.
    addBackgroundImage: (imageId) ->
      model = {
        type: 'background',
        imageId,
      }
      @_media.push(new Image(this, model))

    getBackgroundImageId: ()->
      image = @_media.find(m => m.type == 'background')
      return image && image.imageId
  ###

  set: (col, row, str) ->
    if (arguments.length==1 && col && typeof col == 'object')
      cells = col
      for c,col of cells
        for r,cell of col
          this.set(c,r, cell)


    else if str instanceof Date
      @set col, row, JSDateToExcel str
      # for some reason the number format doesn't apply if the fill is not also set. BUG? Mystery?
      @fill col, row,
        type: "solid",
        fgColor: "FFFFFF"
      @numberFormat col, row, 'd-mmm'
    else if typeof str == 'object'
      for key of str
        @[key] col, row, str[key]
    else if  typeof str == 'string'
      if str != null and str != ''
        @data[row][col].v = @book.ss.str2id('' + str)
      return @data[row][col].dataType = 'string'
    else if typeof str == 'number'
      @data[row][col].v = str
      return @data[row][col].dataType = 'number'
    else
      @data[row][col].v = str
    return



  formula: (col, row, str) ->
    if (typeof str == 'string')
      @formulas = @formulas || []
      @formulas[row] = @formulas[row] || []
      sheet_idx = i for sheet, i in @book.sheets when sheet.name == @name
      @book.cc.add_ref(sheet_idx, col, row)
      @formulas[row][col] = str

  merge: (from_cell, to_cell) ->
    @merges.push({from: from_cell, to: to_cell})

  width: (col, wd) ->
    @col_wd.push({c: col, cw: wd})

  getColWidth: (col) ->
    for _col in @col_wd
      if _col.c == col
        return Math.floor(_col.cw * 10000)
    return 640000

  height: (row, ht) ->
    @row_ht[row] = ht

  font: (col, row, font_s)->
    @styles['font_' + col + '_' + row] = @book.st.font2id(font_s)

  fill: (col, row, fill_s)->
    @styles['fill_' + col + '_' + row] = @book.st.fill2id(fill_s)

  border: (col, row, bder_s)->
    @styles['bder_' + col + '_' + row] = @book.st.bder2id(bder_s)

  numberFormat: (col, row, numfmt_s)->
    @styles['numfmt_' + col + '_' + row] = @book.st.numfmt2id(numfmt_s)

  align: (col, row, align_s)->
    @styles['algn_' + col + '_' + row] = align_s

  valign: (col, row, valign_s)->
    @styles['valgn_' + col + '_' + row] = valign_s

  rotate: (col, row, textRotation)->
    @styles['rotate_' + col + '_' + row] = textRotation

  wrap: (col, row, wrap_s)->
    @styles['wrap_' + col + '_' + row] = wrap_s

  autoFilter: (filter_s) ->
    @autofilter = if typeof filter_s == 'string' then filter_s else @getRange()

  _sheetViews: {
    workbookViewId: '0'
  }

  _sheetViewsPane: {

  }

  _pageSetup: {
    paperSize: '9',
    orientation: 'portrait',
    horizontalDpi: '200',
    verticalDpi: '200'
    }

  sheetViews: (obj) ->
    for key, val of obj
      if (typeof this[key] == 'function')
        this[key](obj[key])
      else
        @_sheetViews[key] = val

  split: (ncols, nrows, state, activePane, topLeftCell) ->
    state = state || "frozen"
    activePane = activePane || "bottomRight"
    topLeftCell = topLeftCell || (tool.i2a((ncols || 0) + 1) + ((nrows || 0) + 1))
    if (ncols)
      @_sheetViewsPane.xSplit = '' + ncols
    if (nrows)
      @_sheetViewsPane.ySplit = '' + nrows
    if (state)
      @_sheetViewsPane.state = state
    if (activePane)
      @_sheetViewsPane.activePane = activePane
    if (topLeftCell)
      @_sheetViewsPane.topLeftCell = topLeftCell

  printBreakRows: (arr) ->
    @_rowBreaks = arr

  printBreakColumns: (arr) ->
    @_colBreaks = arr


  printRepeatRows: (start, end) ->
    if Array.isArray(start)
      @_repeatRows = {start: start[0], end: start[1]}
    else
      @_repeatRows = {start, end}

  printRepeatColumns: (start, end) ->
    if Array.isArray(start)
      @_repeatCols = {start: start[0], end: start[1]}
    else @_repeatCols =  {start, end}

  pageSetup: (obj) ->
    for key, val of obj
      @_pageSetup[key] = val

  pageMargins: (obj) ->
    for key, val of obj
      @_pageMargins[key] = val


  style_id: (col, row) ->
    inx = '_' + col + '_' + row
    style = {
      numfmt_id: @styles['numfmt' + inx],
      font_id: @styles['font' + inx],
      fill_id: @styles['fill' + inx],
      bder_id: @styles['bder' + inx],
      align: @styles['algn' + inx],
      valign: @styles['valgn' + inx],
      rotate: @styles['rotate' + inx],
      wrap: @styles['wrap' + inx]
    }
    id = @book.st.style2id(style)
    return  id

  getRange: () ->
    return '$A$1:$' + tool.i2a(@cols) + '$' + @rows

  toxml: () ->
    ws = xml.create('worksheet', {version: '1.0', encoding: 'UTF-8', standalone: true})
    ws.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    ws.att('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
#    ws.att('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')

#    ws.att('mc:Ignorable', "x14ac xr xr2 xr3")
#    ws.att('xmlns:x14ac', "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
#    ws.att('xmlns:xr', "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
#    ws.att('xmlns:xr2', "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
#    ws.att('xmlns:xr3', "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3")

    ws.ele('dimension', {ref: 'A1'})

    ws.ele('sheetViews').ele('sheetView', @_sheetViews).ele('pane', @_sheetViewsPane)


    ws.ele('sheetFormatPr', {defaultRowHeight: '13.5'})
    if @col_wd.length > 0
      cols = ws.ele('cols')
      for cw in @col_wd
        cols.ele('col', {min: '' + cw.c, max: '' + cw.c, width: cw.cw, customWidth: '1'})
    sd = ws.ele('sheetData')
    for i in [1..@rows]
      r = sd.ele('row', {r: '' + i, spans: '1:' + @cols})
      ht = @row_ht[i]
      if ht
        r.att('ht', ht)
        r.att('customHeight', '1')
      for j in [1..@cols]
        ix = @data[i][j]
        sid = @style_id(j, i)
        if (ix.v isnt null and ix.v isnt undefined) or (sid isnt 1)
          c = r.ele('c', {r: '' + tool.i2a(j) + i})
          c.att('s', '' + (sid - 1)) if sid isnt 1
          if (@formulas[i] && @formulas[i][j])
            c.ele('f', '' + @formulas[i][j])
            c.ele('v')
          else if ix.dataType == 'string'
            c.att('t', 's')
            c.ele('v', '' + (ix.v - 1))
          else if ix.dataType == 'number'
            c.ele 'v', '' + ix.v



    if @merges.length > 0
      mc = ws.ele('mergeCells', {count: @merges.length})
      for m in @merges
        mc.ele('mergeCell', {ref: ('' + tool.i2a(m.from.col) + m.from.row + ':' + tool.i2a(m.to.col) + m.to.row)})
    if typeof @autofilter == 'string'
      ws.ele('autoFilter', {ref: @autofilter})
    ws.ele('phoneticPr', {fontId: '1', type: 'noConversion'})

    ws.ele('pageMargins', @_pageMargins)
    ws.ele('pageSetup', @_pageSetup)

    if @_rowBreaks && @_rowBreaks.length
      cb = ws.ele('rowBreaks', {count: @_rowBreaks.length, manualBreakCount: @_rowBreaks.length})
      for i in @_rowBreaks
        cb.ele('brk', { id: i, man: '1'})

    if @_colBreaks && @_colBreaks.length
      cb = ws.ele('colBreaks', {count: @_colBreaks.length, manualBreakCount: @_colBreaks.length})
      for i in @_colBreaks
        cb.ele('brk', { id: i, man: '1'})

    for wsRel in @wsRels
      ws.ele('drawing', {'r:id': wsRel.id})

    ws.end({pretty: false})

