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

