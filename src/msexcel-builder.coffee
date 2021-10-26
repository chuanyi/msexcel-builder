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

class XlWorksheetRels
  constructor: (@wsRels) ->
  generate: () ->

  toxml: () ->
    rs = xml.create('Relationships', {version: '1.0', encoding: 'UTF-8', standalone: true})
    rs.att('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships')
    for wsRel in @wsRels
      rs.ele('Relationship', {
        Id: wsRel.id,
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
        Target: wsRel.target
      })
    rs.end({pretty: false})

class XlDrawingRels
  constructor: (@dwRels) ->
  generate: () ->

  toxml: () ->
    rs = xml.create('Relationships', {version: '1.0', encoding: 'UTF-8', standalone: true})
    rs.att('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships')
    for dwRel in @dwRels
      rs.ele('Relationship', {
        Id: dwRel.id,
# Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        Target: dwRel.target
      })
    rs.end({pretty: false})

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
    if str instanceof Date
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

class CalcChain

  constructor: (@book)->
    @cache = {}

  add_ref: (idx, col, row)->
    num = idx + 1
    if !@cache.hasOwnProperty(num)
      @cache[num] = []
    @cache[num].push(tool.i2a(col) + row)

  toxml: ()->
    cc = xml.create('calcChain', {version: '1.0', encoding: 'UTF-8', standalone: true})
    cc.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

    for key, val of @cache
      for el in val
        cc.ele('c', {r: '' + el, i: '' + key})

    return cc.end({pretty: false})

class CalcChain

  constructor: (@book)->
    @cache = {}

  add_ref: (idx, col, row)->
    num = idx + 1
    if !@cache.hasOwnProperty(num)
      @cache[num] = []
    @cache[num].push(tool.i2a(col) + row)

  toxml: ()->
    cc = xml.create('calcChain', {version: '1.0', encoding: 'UTF-8', standalone: true})
    cc.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

    for key, val of @cache
      for el in val
        cc.ele('c', { r: '' + el, i: '' + key })

    return cc.end()

class Style

  numberFormats: {
    0: 'General'
    1: '0'
    2: '0.00'
    3: '#,##0'
    4: '#,##0.00'
    9: '0%'
    10: '0.00%'
    11: '0.00E+00'
    12: '# ?/?'
    13: '# ??/??'
    14: 'm/d/yy'
    15: 'd-mmm-yy'
    16: 'd-mmm'
    17: 'mmm-yy'
    18: 'h:mm AM/PM'
    19: 'h:mm:ss AM/PM'
    20: 'h:mm'
    21: 'h:mm:ss'
    22: 'm/d/yy h:mm'
    37: '#,##0 ;(#,##0)'
    38: '#,##0 ;[Red](#,##0)'
    39: '#,##0.00;(#,##0.00)'
    40: '#,##0.00;[Red](#,##0.00)'
    45: 'mm:ss'
    46: '[h]:mm:ss'
    47: 'mmss.0'
    48: '##0.0E+0'
    49: '@'
    56: '"上午/下午 "hh"時"mm"分"ss"秒 "'
  }

  constructor: (@book)->
    @cache = {}
    @mfonts = [] # font style
    @mfills = [] # fill style
    @mbders = [] # border style
    @mstyle = [] # cell style<ref-font,ref-fill,ref-border,align>
    @numFmtNextId = 164
    @with_default()

  with_default: ()->
    @def_font_id = @font2id(null)
    @def_fill_id = @fill2id(null)
    @def_bder_id = @bder2id(null)
    @def_align = '-'
    @def_valign = '-'
    @def_rotate = '-'
    @def_wrap = '-'
    @def_numfmt_id = 0
    @def_style_id = @style2id({
      font_id: @def_font_id,
      fill_id: @def_fill_id,
      bder_id: @def_bder_id,
      align: @def_align,
      valign: @def_valign,
      rotate: @def_rotate
    })

  font2id: (font)->
    font or= {}
    font.bold or= '-'
    font.iter or= '-'
    font.sz or= '11'
    font.color or= '-'
    font.name or= 'Calibri'
    font.scheme or= 'minor'
    font.family or= '2'

    font.underline or= '-'
    font.strike or= '-'
    font.outline or= '-'
    font.shadow or= '-'

    k = 'font_' + font.bold + font.iter + font.sz + font.color + font.name + font.scheme + font.family + font.underline + font.strike + font.outline + font.shadow

    id = @cache[k]
    if id
      return id
    else
      @mfonts.push font
      @cache[k] = @mfonts.length
      return @mfonts.length

  fill2id: (fill)->
    fill or= {}
    fill.type or= 'none'
    fill.bgColor or= '-'
    fill.fgColor or= '-'
    k = 'fill_' + fill.type + fill.bgColor + fill.fgColor
    id = @cache[k]
    if id
      return id
    else
      @mfills.push fill
      @cache[k] = @mfills.length
      return @mfills.length

  bder2id: (bder)->
    bder or= {}
    bder.left or= '-'
    bder.right or= '-'
    bder.top or= '-'
    bder.bottom or= '-'
    k = JSON.stringify(["bder_",bder.left,bder.right, bder.top, bder.bottom])
    id = @cache[k]
    if id
      return id
    else
      @mbders.push bder
      @cache[k] = @mbders.length
      return @mbders.length

  numfmt2id: (numfmt) ->
    if typeof numfmt == 'number'
      return numfmt
    else if typeof numfmt == 'string'
      for key of @numberFormats
        if @numberFormats[key] == numfmt
          return parseInt key;
      # if it's not in numberFormats, we parse the string and add it the end of numberFormats
      if !numfmt
        throw "Invalid format specification"
      numfmt = numfmt
        .replace(/&/g, '&amp')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
      @numberFormats[++@numFmtNextId] = numfmt
      return @numFmtNextId

  style2id: (style)->
    style.align or= @def_align
    style.valign or= @def_valign
    style.rotate or= @def_rotate
    style.wrap or= @def_wrap
    style.font_id or= @def_font_id
    style.fill_id or= @def_fill_id
    style.bder_id or= @def_bder_id
    style.numfmt_id or= @def_numfmt_id
    k = 's_' + [style.font_id, style.fill_id, style.bder_id, style.align, style.valign, style.wrap, style.rotate,
      style.numfmt_id].join('_')
    id = @cache[k]
    if id
      return id
    else
      @mstyle.push style
      @cache[k] = @mstyle.length
      return @mstyle.length

  toxml: ()->
    ss = xml.create('styleSheet', {version: '1.0', encoding: 'UTF-8', standalone: true})
    ss.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    # add all numFmts >= 164 as <numFmt numFmtId="${o.num_fmt_id}" formatCode="numFmt"/>
    customNumFmts = [];
    for key, fmt of @numberFormats
      if parseInt(key) >= 164
        customNumFmts.push({numFmtId: key, formatCode: fmt});
    if customNumFmts.length > 0
      numFmts = ss.ele('numFmts', {
        count: customNumFmts.length
      });
      for o in customNumFmts
        numFmts.ele('numFmt', o)
    fonts = ss.ele('fonts', {count: @mfonts.length})
    for o in @mfonts
      e = fonts.ele('font')
      e.ele('b') if o.bold isnt '-'
      e.ele('i') if o.iter isnt '-'
      e.ele('u') if o.underline isnt '-'
      e.ele('strike') if o.strike isnt '-'
      e.ele('outline') if o.outline isnt '-'
      e.ele('shadow') if o.shadow isnt '-'

      e.ele('sz', {val: o.sz})
      e.ele('color', {rgb: o.color}) if o.color isnt '-'
      e.ele('name', {val: o.name})
      e.ele('family', {val: o.family})
      e.ele('charset', {val: '134'})
      e.ele('scheme', {val: 'minor'}) if o.scheme isnt '-'
    fills = ss.ele('fills', {count: @mfills.length + 2})
    fills.ele('fill').ele('patternFill', {patternType: 'none'})
    fills.ele('fill').ele('patternFill', {patternType: 'gray125'})
    #<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>

    for o in @mfills
      e = fills.ele('fill')
      es = e.ele('patternFill', {patternType: o.type})
      es.ele('fgColor', {rgb: o.fgColor}) if o.fgColor isnt '-'
      es.ele('bgColor', {indexed: o.bgColor}) if o.bgColor isnt '-'
    bders = ss.ele('borders', {count: @mbders.length})
    for o in @mbders

      e = bders.ele('border')

      if o.left isnt '-'
        if typeof o.left is 'string'
          e.ele('left', {style: o.left}).ele('color', {auto: '1'})
        else
          e.ele('left', {style: o.left.style}).ele('color', o.left.color)
      else e.ele('left')

      if o.right isnt '-'
        if typeof o.right is 'string'
          e.ele('right', {style: o.right}).ele('color', {auto: '1'})
        else
          e.ele('right', {style: o.right.style}).ele('color', o.right.color)
      else e.ele('right')

      if o.top isnt '-'
        if typeof o.top is 'string'
          e.ele('top', {style: o.top}).ele('color', {auto: '1'})
        else
          e.ele('top', {style: o.top.style}).ele('color', o.top.color)
      else e.ele('top')

      if o.bottom isnt '-'
        if typeof o.bottom is 'string'
          e.ele('bottom', {style: o.bottom}).ele('color', {auto: '1'})
        else
          e.ele('bottom', {style: o.bottom.style}).ele('color', o.bottom.color)
      else e.ele('bottom')



      e.ele('diagonal')
    ss.ele('cellStyleXfs', {count: '1'}).ele('xf', {
      numFmtId: '0',
      fontId: '0',
      fillId: '0',
      borderId: '0'
    }).ele('alignment', {vertical: 'center'})
    cs = ss.ele('cellXfs', {count: @mstyle.length})
    for o in @mstyle
      e = cs.ele('xf', {
        numFmtId: o.numfmt_id || '0',
        fontId: (o.font_id - 1),
        fillId: o.fill_id + 1,
        borderId: (o.bder_id - 1),
        xfId: '0'
      })
      e.att('applyFont', '1') if o.font_id isnt 1
      e.att('applyFill', '1') if o.fill_id isnt 1
      e.att('applyNumberFormat', '1') if o.numfmt_id isnt undefined
      e.att('applyBorder', '1') if o.bder_id isnt 1
      if o.align isnt '-' or o.valign isnt '-' or o.wrap isnt '-'
        e.att('applyAlignment', '1')
        ea = e.ele('alignment', {
          textRotation: (if o.rotate is '-' then '0' else o.rotate),
          horizontal: (if o.align is '-' then 'left' else o.align),
          vertical: (if o.valign is '-' then 'bottom' else o.valign)
        })
        ea.att('wrapText', '1') if o.wrap isnt '-'
    ss.ele('cellStyles', {count: '1'}).ele('cellStyle', {name: 'Normal', xfId: '0', builtinId: '0'})
    ss.ele('dxfs', {count: '0'})
    ss.ele('tableStyles', {count: '0', defaultTableStyle: 'TableStyleMedium9', defaultPivotStyle: 'PivotStyleLight16'})
    return ss.end({pretty: false})

class Image
  constructor: (@id, @extension, @content, @range, @options) ->
    @editAs = 'oneCell'


  publish: (sheet, zip, media)->
    ## Inject image, it's used by sheet
    ## 1. write data to media folder
    ##    convert base 64 to text
    ##    define filename based on number of existing medias
    ##    writes to media
    ##
    ## 2. create reference for media to drawing
    ## 3. create the actual drawing using reference for media and set location
    ## 4. creates reference for drawing to sheet.
    ## 5. use image rel to sheet.

  toDrawingXml: (relId,spec)->
    # pngVersionRel = 'rId1'
    # svgVersionRel = 'rId2'

    dr = xml.create('xdr:wsDr', {version: '1.0', encoding: 'UTF-8', standalone: true})
    # dr.att('xmlns:xdr', 'http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing')
    dr.att('xmlns:xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing')
    # dr.att('xmlns:a', 'http://purl.oclc.org/ooxml/drawingml/main')
    dr.att('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main')

    twoCellAnchor = dr.ele('xdr:twoCellAnchor', {editAs: @editAs})

    _from = twoCellAnchor.ele('xdr:from') # {col: 0, colOff: 0, row: 1, rowOff: 0})
    _from.ele('xdr:col', @range.from.model.nativeCol)
    _from.ele('xdr:colOff', @range.from.model.nativeColOff)
    _from.ele('xdr:row', @range.from.model.nativeRow)
    _from.ele('xdr:rowOff', @range.from.model.nativeRowOff)

    _to = twoCellAnchor.ele('xdr:to') # {col: 5, colOff: 419100, row: 16, rowOff: 0})
    _to.ele('xdr:col', @range.to.model.nativeCol)
    _to.ele('xdr:colOff', @range.to.model.nativeColOff)
    _to.ele('xdr:row', @range.to.model.nativeRow)
    _to.ele('xdr:rowOff', @range.to.model.nativeRowOff)

    pic = twoCellAnchor.ele('xdr:pic')
    nvPicPr = pic.ele('xdr:nvPicPr')
    ## Graphic name printed on xml are not seen or used as reference
    ## Skiping proper calculation part and just finding a random until 100000
    graphic_index = 30924 || Math.round(Math.random() * 100000)
    cNvPr = nvPicPr.ele('xdr:cNvPr', {id: 3, name: 'Graphic ' + graphic_index})
    cNvPr.ele('a:extLst')
      .ele('a:ext', {uri: '{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}'})
      .ele('a16:creationId', {
      'xmlns:a16': 'http://schemas.microsoft.com/office/drawing/2014/main'
    , 'id': '{9D66B5F7-2553-484C-A5BE-4D0B8D57E08B}'
    })
    nvPicPr.ele('xdr:cNvPicPr')
      .ele('a:picLocks', {noChangeAspect: 1})

    blipFill = pic.ele('xdr:blipFill')


    blipModel = {
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }

    if (relId)
      blipModel["r:embed"]=relId

    blipModel['cstate'] = "print"
    blip = blipFill.ele('a:blip', blipModel)
    extLst = blip.ele('a:extLst')
    ext = extLst.ele('a:ext', {
      'uri': '{28A0092B-C50C-407E-A947-70E740481C1C}'
    })
    ext.ele('a14:useLocalDpi', {
      'xmlns:a14': 'http://schemas.microsoft.com/office/drawing/2010/main'
    , 'val': 0
    })

    if @extension =='svg'
      ext = extLst.ele('a:ext', {
        'uri': '{96DAC541-7B7A-43D3-8B79-37D633B846F1}'
      })
      ext.ele('asvg:svgBlip', {
        'xmlns:asvg':'http://schemas.microsoft.com/office/drawing/2016/SVG/main'
        , 'r:embed': relId
      })


    if @options.stretch
      blipFill.ele('a:stretch')
    blipFill.ele('srcRect')
    #.ele('a:fillRect')
    spPr = pic.ele('xdr:spPr')
    xfrm = spPr.ele('a:xfrm')
    xfrm.ele('a:off', {'x': 609600, 'y': 190500})
    xfrm.ele('a:ext', {'cx': 2857500, 'cy': 2857500})
    spPr.ele('a:prstGeom', {'prst': 'rect'}).ele('a:avLst')
    twoCellAnchor.ele('xdr:clientData')
    dr.end({pretty: false})

class Workbook
  constructor: (@fpath, @fname) ->
    @id = '' + parseInt(Math.random() * 9999999)
    # create temp folder & copy template data
    # init
    @sheets = []
    @medias = []
    @ss = new SharedStrings
    @ct = new ContentTypes(@)
    @da = new DocPropsApp(@)
    @wb = new XlWorkbook(@)
    @wbre = new XlWorkbookRels(@)
    @st = new Style(@)
    @cc = new CalcChain(@)
    @cc = new CalcChain(@)
# @wsre = new XlSheetRels(@)
# @dw = new XlDrawing(@)
# @dwre = new XlDrawingRels(@)

  createSheet: (name, cols, rows) ->
    sheet = new Sheet(@, name, cols, rows)
    @sheets.push sheet
    return sheet

  _addMediaFromImage: (image) ->
## converts image into proper media data structure
    @medias.push({image: image})

  _removeMediaFromImage: (image) ->
    foundIndex = -1
    ## find image.
    for media, i in @medias
      if media.image.id = image.id
        foundIndex = i
        break
    ## remove it if found.
    @medias.splice(foundIndex, 1) unless foundIndex > -1

# _addDrawingFromImage: (image, media) ->
##
## @dwMediaRel = {}
## @dwMediaRels.push(dwMediaRel)
## @dw = {}
## @dws.push(dw)

  save: (target, opts, cb) ->
    if (arguments.length == 1 && typeof target == 'function')
      cb = target
      target = @fpath + '/' + @fname
      opts = {}
    else if (arguments.length == 2 && typeof opts =='function')
      cb = opts
      opts = {}

    @generate (err, zip) ->
      buffer = undefined
      args = {type: 'nodebuffer'}
      if (opts.compressed)
        args.compressed = "DEFLATE"

      buffer = zip.generateAsync(args).then((buffer) ->
        if err
          return cb(err)
        fs.writeFile target, buffer, cb
      )

# takes a callback function(err, zip) and returns a JSZip object on success
  generate: (cb) =>
    zip = new JSZip()

    for key of baseXl
      zip.file key, baseXl[key]

    # 1 - build [Content_Types].xml
    zip.file('[Content_Types].xml', @ct.toxml())
    # 2 - build docProps/app.xml
    zip.file('docProps/app.xml', @da.toxml())
    # 3 - build xl/workbook.xml
    zip.file('xl/workbook.xml', @wb.toxml())
    # 4 - build xl/sharedStrings.xml
    zip.file('xl/sharedStrings.xml', @ss.toxml())
    # 5 - build xl/_rels/workbook.xml.rels
    zip.file('xl/_rels/workbook.xml.rels', @wbre.toxml())

    wbMediaCounter = 1 # workbook media counter, per generation

    for sheet, i in @sheets
      sheet.wsRels = []

      for image, j in sheet.images

        dwRels = [] # drawing media list, for this release, one per drawing
        relId = 'rId' + (sheet.wsRels.length + 1) # same for both media to drawing OR sheet to drawing rels.

        # - build xl/media/image(1-N).xml
        mediaFilename = [wbMediaCounter, '.', image.extension].join('')
        zip.file('xl/media/image' + mediaFilename, image.content, {base64: true})
        dwRels.push({id: relId, target: '../media/image' + mediaFilename})

        # - build xl/drawings/drawing(1-N).xml
        drawingFilename = wbMediaCounter + '.xml'
        sheet.wsRels.push({id: relId, target: '../drawings/drawing' + drawingFilename})
        zip.file('xl/drawings/drawing' + drawingFilename, image.toDrawingXml(relId, image))

        # - build xl/drawings/_rels/drawing(1-N).xml.rels
        zip.file('xl/drawings/_rels/drawing' + wbMediaCounter + '.xml.rels', new XlDrawingRels(dwRels).toxml())

        wbMediaCounter++

      # - build xl/worksheets/_rels/sheet(1-N).xml.rels
      zip.file('xl/worksheets/_rels/sheet' + (i + 1) + '.xml.rels', new XlWorksheetRels(sheet.wsRels).toxml())

      # - build xl/worksheets/sheet(1-N).xml
      zip.file('xl/worksheets/sheet' + (i + 1) + '.xml', @sheets[i].toxml())
    # 7 - build xl/styles.xml
    zip.file('xl/styles.xml', @st.toxml())
    # 8 - build xl/calcChain.xml
    if Object.getOwnPropertyNames(@cc.cache).length > 0
      zip.file('xl/calcChain.xml', @cc.toxml())
    # 9 - build xl/worksheets/sheet(1-N).xml

    cb(null, zip)

  cancel: () ->
# delete temp folder
    console.error "workbook.cancel() is deprecated"

JSDateToExcel = (dt) ->
  dt.valueOf() / 86400000 + 25569

if (module? && module.exports?)
  module.exports =
    createWorkbook: (fpath, fname)->
      return new Workbook(fpath, fname)

if (window?)
  window.excelbuilder =
    createWorkbook: (fpath, fname)->
      return new Workbook(fpath, fname)

# Base content formerly stored in /lib/tmpl but placed in code so as to avoid dependence on file system
baseXl =
  '_rels/.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
  'docProps/core.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Administrator</dc:creator><cp:lastModifiedBy></cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2006-09-13T11:21:51Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2006-09-13T11:21:55Z</dcterms:modified></cp:coreProperties>'
  'xl/theme/theme1.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office 主题"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'
  'xl/styles.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><color theme="1"/><name val="宋体"/><family val="2"/><charset val="134"/><scheme val="minor"/></font><font><sz val="9"/><name val="宋体"/><family val="2"/><charset val="134"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/></styleSheet>'
