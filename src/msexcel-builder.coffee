###
  MS Excel 2007 Creater v0.0.1
  Author : chuanyi.zheng@gmail.com
  History: 2012/11/07 first created
###
JSZip = require 'jszip'
xml = require 'xmlbuilder'

tool =
  i2a : (i) ->
    return 'ABCDEFGHIJKLMNOPQRSTUVWXYZ123'.charAt(i-1)

opt = 
  tmpl_path : __dirname

class ContentTypes
  constructor: (@book)->

  toxml:()->
    types = xml.create('Types',{version:'1.0',encoding:'UTF-8',standalone:true})
    types.att('xmlns','http://schemas.openxmlformats.org/package/2006/content-types')
    types.ele('Override',{PartName:'/xl/theme/theme1.xml',ContentType:'application/vnd.openxmlformats-officedocument.theme+xml'})
    types.ele('Override',{PartName:'/xl/styles.xml',ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'})
    types.ele('Default',{Extension:'rels',ContentType:'application/vnd.openxmlformats-package.relationships+xml'})
    types.ele('Default',{Extension:'xml',ContentType:'application/xml'})
    types.ele('Override',{PartName:'/xl/workbook.xml',ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'})
    types.ele('Override',{PartName:'/docProps/app.xml',ContentType:'application/vnd.openxmlformats-officedocument.extended-properties+xml'})
    for i in [1..@book.sheets.length]
      types.ele('Override',{PartName:'/xl/worksheets/sheet'+i+'.xml',ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'})
    types.ele('Override',{PartName:'/xl/sharedStrings.xml',ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'})
    types.ele('Override',{PartName:'/docProps/core.xml',ContentType:'application/vnd.openxmlformats-package.core-properties+xml'})
    return types.end()

class DocPropsApp
  constructor: (@book)->

  toxml: ()->
    props = xml.create('Properties',{version:'1.0',encoding:'UTF-8',standalone:true})
    props.att('xmlns','http://schemas.openxmlformats.org/officeDocument/2006/extended-properties')
    props.att('xmlns:vt','http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes')
    props.ele('Application','Microsoft Excel')
    props.ele('DocSecurity','0')
    props.ele('ScaleCrop','false')
    tmp = props.ele('HeadingPairs').ele('vt:vector',{size:2,baseType:'variant'})
    tmp.ele('vt:variant').ele('vt:lpstr','工作表')
    tmp.ele('vt:variant').ele('vt:i4',''+@book.sheets.length)
    tmp = props.ele('TitlesOfParts').ele('vt:vector',{size:@book.sheets.length,baseType:'lpstr'})
    for i in [1..@book.sheets.length]
      tmp.ele('vt:lpstr',@book.sheets[i-1].name)
    props.ele('Company')
    props.ele('LinksUpToDate','false')
    props.ele('SharedDoc','false')  
    props.ele('HyperlinksChanged','false')  
    props.ele('AppVersion','12.0000') 
    return props.end()

class XlWorkbook
  constructor: (@book)->

  toxml: ()->
    wb = xml.create('workbook',{version:'1.0',encoding:'UTF-8',standalone:true})
    wb.att('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    wb.att('xmlns:r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    wb.ele('fileVersion ',{appName:'xl',lastEdited:'4',lowestEdited:'4',rupBuild:'4505'})
    wb.ele('workbookPr',{filterPrivacy:'1',defaultThemeVersion:'124226'}) 
    wb.ele('bookViews').ele('workbookView ',{xWindow:'0',yWindow:'90',windowWidth:'19200',windowHeight:'11640'})
    tmp = wb.ele('sheets')
    for i in [1..@book.sheets.length]
      tmp.ele('sheet',{name:@book.sheets[i-1].name,sheetId:''+i,'r:id':'rId'+i})
    wb.ele('calcPr',{calcId:'124519'})
    return wb.end()

class XlRels
  constructor: (@book)->
  
  toxml: ()->
    rs = xml.create('Relationships',{version:'1.0',encoding:'UTF-8',standalone:true})
    rs.att('xmlns','http://schemas.openxmlformats.org/package/2006/relationships')
    for i in [1..@book.sheets.length]
      rs.ele('Relationship',{Id:'rId'+i,Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',Target:'worksheets/sheet'+i+'.xml'})
    rs.ele('Relationship',{Id:'rId'+(@book.sheets.length+1),Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',Target:'theme/theme1.xml'})
    rs.ele('Relationship',{Id:'rId'+(@book.sheets.length+2),Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',Target:'styles.xml'})
    rs.ele('Relationship',{Id:'rId'+(@book.sheets.length+3),Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',Target:'sharedStrings.xml'})
    return rs.end()

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
    sst = xml.create('sst',{version:'1.0',encoding:'UTF-8',standalone:true})
    sst.att('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    sst.att('count',''+@arr.length)
    sst.att('uniqueCount',''+@arr.length)
    for i in [0...@arr.length]
      si = sst.ele('si')
      si.ele('t',@arr[i])
      si.ele('phoneticPr',{fontId:1,type:'noConversion'})
    return sst.end()

class Sheet
  constructor: (@book, @name, @cols, @rows) ->
    @data = {}
    for i in [1..@rows]
      @data[i] = {}
      for j in [1..@cols]
        @data[i][j] = {v:0}
    @merges = []
    @col_wd = []
    @row_ht = {}
    @styles = {}

  set: (col, row, str) ->
    @data[row][col].v = @book.ss.str2id(''+str) if str? and str isnt ''

  merge: (from_cell, to_cell) ->
    @merges.push({from:from_cell, to:to_cell})

  width: (col, wd) ->
    @col_wd.push {c:col,cw:wd}

  height: (row, ht) ->
    @row_ht[row] = ht

  font: (col, row, font_s)->
    @styles['font_'+col+'_'+row] = @book.st.font2id(font_s)

  fill: (col, row, fill_s)-> 
    @styles['fill_'+col+'_'+row] = @book.st.fill2id(fill_s)

  border: (col, row, bder_s)->
    @styles['bder_'+col+'_'+row] = @book.st.bder2id(bder_s)

  align: (col, row, align_s)->
    @styles['algn_'+col+'_'+row] = align_s

  valign: (col, row, valign_s)->
    @styles['valgn_'+col+'_'+row] = valign_s

  rotate: (col, row, textRotation)->
    @styles['rotate_'+col+'_'+row] = textRotation

  wrap: (col, row, wrap_s)->
    @styles['wrap_'+col+'_'+row] = wrap_s

  style_id: (col, row) ->
    inx = '_'+col+'_'+row
    style = {font_id:@styles['font'+inx],fill_id:@styles['fill'+inx],bder_id:@styles['bder'+inx],align:@styles['algn'+inx],valign:@styles['valgn'+inx],rotate:@styles['rotate'+inx],wrap:@styles['wrap'+inx]}
    id = @book.st.style2id(style)
    return id

  toxml: () ->
    ws = xml.create('worksheet',{version:'1.0',encoding:'UTF-8',standalone:true})
    ws.att('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    ws.att('xmlns:r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    ws.ele('dimension',{ref:'A1'})
    ws.ele('sheetViews').ele('sheetView',{workbookViewId:'0'})
    ws.ele('sheetFormatPr',{defaultRowHeight:'13.5'})
    if @col_wd.length > 0
      cols = ws.ele('cols')
      for cw in @col_wd
        cols.ele('col',{min:''+cw.c,max:''+cw.c,width:cw.cw,customWidth:'1'})
    sd = ws.ele('sheetData')
    for i in [1..@rows]
      r = sd.ele('row',{r:''+i,spans:'1:'+@cols})
      ht = @row_ht[i]
      if ht
        r.att('ht',ht)
        r.att('customHeight','1')
      for j in [1..@cols]
        ix = @data[i][j]
        sid = @style_id(j, i)
        if (ix.v isnt 0) or (sid isnt 1)
          c = r.ele('c',{r:''+tool.i2a(j)+i})
          c.att('s',''+(sid-1)) if sid isnt 1
          if ix.v isnt 0
            c.att('t','s')
            c.ele('v',''+(ix.v-1))
    if @merges.length > 0
      mc = ws.ele('mergeCells',{count:@merges.length})
      for m in @merges
        mc.ele('mergeCell',{ref:(''+tool.i2a(m.from.col)+m.from.row+':'+tool.i2a(m.to.col)+m.to.row)})
    ws.ele('phoneticPr',{fontId:'1',type:'noConversion'})
    ws.ele('pageMargins',{left:'0.7',right:'0.7',top:'0.75',bottom:'0.75',header:'0.3',footer:'0.3'})
    ws.ele('pageSetup',{paperSize:'9',orientation:'portrait',horizontalDpi:'200',verticalDpi:'200'})
    return ws.end()

class Style
  constructor: (@book)->
    @cache = {}
    @mfonts = []  # font style
    @mfills = []  # fill style
    @mbders = []  # border style
    @mstyle = []  # cell style<ref-font,ref-fill,ref-border,align>
    @with_default()

  with_default:()->
    @def_font_id = @font2id(null)
    @def_fill_id = @fill2id(null)
    @def_bder_id = @bder2id(null)
    @def_align = '-'
    @def_valign = '-'
    @def_rotate = '-'
    @def_wrap = '-'
    @def_style_id = @style2id({font_id:@def_font_id,fill_id:@def_fill_id,bder_id:@def_bder_id,align:@def_align,valign:@def_valign,rotate:@def_rotate})

  font2id: (font)->
    font or= {}
    font.bold or= '-'
    font.iter or= '-'
    font.sz or= '11'
    font.color or= '-'
    font.name or= '宋体'
    font.scheme or='minor'
    font.family or= '2'
    k = 'font_'+font.bold+font.iter+font.sz+font.color+font.name+font.scheme+font.family
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
    k = 'bder_'+bder.left+'_'+bder.right+'_'+bder.top+'_'+bder.bottom
    id = @cache[k]
    if id
      return id
    else
      @mbders.push bder
      @cache[k] = @mbders.length
      return @mbders.length

  style2id:(style)->
    style.align or= @def_align
    style.valign or= @def_valign
    style.rotate or= @def_rotate
    style.wrap or= @def_wrap
    style.font_id or= @def_font_id
    style.fill_id or= @def_fill_id
    style.bder_id or= @def_bder_id
    k = 's_' + style.font_id + '_' + style.fill_id + '_' + style.bder_id + '_' + style.align + '_' + style.valign + '_' + style.wrap + '_' + style.rotate
    id = @cache[k]
    if id
      return id
    else
      @mstyle.push style
      @cache[k] = @mstyle.length
      return @mstyle.length

  toxml: ()->
    ss = xml.create('styleSheet',{version:'1.0',encoding:'UTF-8',standalone:true})
    ss.att('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    fonts = ss.ele('fonts',{count:@mfonts.length})
    for o in @mfonts
      e = fonts.ele('font')
      e.ele('b') if o.bold isnt '-'
      e.ele('i') if o.iter isnt '-'
      e.ele('sz',{val:o.sz})
      e.ele('color',{theme:o.color}) if o.color isnt '-'
      e.ele('name',{val:o.name})
      e.ele('family',{val:o.family})
      e.ele('charset',{val:'134'})
      e.ele('scheme',{val:'minor'}) if o.scheme isnt '-'
    fills = ss.ele('fills',{count:@mfills.length})
    for o in @mfills
      e = fills.ele('fill')
      es = e.ele('patternFill',{patternType:o.type})
      es.ele('fgColor',{theme:'8',tint:'0.79998168889431442'}) if o.fgColor isnt '-'
      es.ele('bgColor',{indexed:o.bgColor}) if o.bgColor isnt '-'
    bders = ss.ele('borders',{count:@mbders.length})
    for o in @mbders
      e = bders.ele('border')
      if o.left isnt '-' then e.ele('left',{style:o.left}).ele('color',{auto:'1'}) else e.ele('left')
      if o.right isnt '-' then e.ele('right',{style:o.right}).ele('color',{auto:'1'}) else e.ele('right')
      if o.top isnt '-' then e.ele('top',{style:o.top}).ele('color',{auto:'1'}) else e.ele('top')
      if o.bottom isnt '-' then e.ele('bottom',{style:o.bottom}).ele('color',{auto:'1'}) else e.ele('bottom')
      e.ele('diagonal')
    ss.ele('cellStyleXfs',{count:'1'}).ele('xf',{numFmtId:'0',fontId:'0',fillId:'0',borderId:'0'}).ele('alignment',{vertical:'center'})
    cs = ss.ele('cellXfs',{count:@mstyle.length})
    for o in @mstyle
      e = cs.ele('xf',{numFmtId:'0',fontId:(o.font_id-1),fillId:(o.fill_id-1),borderId:(o.bder_id-1),xfId:'0'})
      e.att('applyFont','1') if o.font_id isnt 1
      e.att('applyFill','1') if o.fill_id isnt 1
      e.att('applyBorder','1') if o.bder_id isnt 1
      if o.align isnt '-' or o.valign isnt '-' or o.wrap isnt '-'
        e.att('applyAlignment','1')
        ea = e.ele('alignment',{textRotation:(if o.rotate is '-' then '0' else o.rotate),horizontal:(if o.align is '-' then 'left' else o.align), vertical:(if o.valign is '-' then 'top' else o.valign)})
        ea.att('wrapText','1') if o.wrap isnt '-'
    ss.ele('cellStyles',{count:'1'}).ele('cellStyle',{name:'常规',xfId:'0',builtinId:'0'})
    ss.ele('dxfs',{count:'0'})
    ss.ele('tableStyles',{count:'0',defaultTableStyle:'TableStyleMedium9',defaultPivotStyle:'PivotStyleLight16'})
    return ss.end()

class Workbook
  constructor: (@fpath, @fname) ->
    @id = ''+parseInt(Math.random()*9999999)
    # create temp folder & copy template data
    # init
    @sheets = []
    @ss = new SharedStrings
    @ct = new ContentTypes(@)
    @da = new DocPropsApp(@)
    @wb = new XlWorkbook(@)
    @re = new XlRels(@)
    @st = new Style(@)

  createSheet: (name, cols, rows) ->
    sheet = new Sheet(@,name,cols,rows)
    @sheets.push sheet
    return sheet

  save: (cb) =>
    target = @fpath + '/' + @fname
    @generate (err, zip) ->
      buffer = zip.generate(type: 'nodebuffer')
      # dependence on file system isolated to this function
      require('fs').writeFile target, buffer, cb

 # takes a callback function(err, zip) and returns a JSZip object on success
  generate: (cb) =>

    zip = new JSZip

    for key of baseXl
      zip.file key, baseXl[key]

    # 1 - build [Content_Types].xml
    zip.file('[Content_Types].xml',@ct.toxml())
    # 2 - build docProps/app.xml
    zip.file('docProps/app.xml',@da.toxml())
    # 3 - build xl/workbook.xml
    zip.file('xl/workbook.xml',@wb.toxml())
    # 4 - build xl/sharedStrings.xml
    zip.file('xl/sharedStrings.xml',@ss.toxml())
    # 5 - build xl/_rels/workbook.xml.rels
    zip.file('xl/_rels/workbook.xml.rels',@re.toxml())
    # 6 - build xl/worksheets/sheet(1-N).xml
    for i in [0...@sheets.length]
      zip.file('xl/worksheets/sheet'+(i+1)+'.xml',@sheets[i].toxml())
    # 7 - build xl/styles.xml
    zip.file('xl/styles.xml',@st.toxml())
    cb null, zip

  cancel: () ->
    # delete temp folder
    console.error "workbook.cancel() is deprecated"

module.exports =
  createWorkbook: (fpath, fname)->
    return new Workbook(fpath, fname)

# Base content formerly stored in /lib/tmpl but placed in code so as to avoid dependence on file system
baseXl =
  '_rels/.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
  'docProps/core.xml':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Administrator</dc:creator><cp:lastModifiedBy></cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2006-09-13T11:21:51Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2006-09-13T11:21:55Z</dcterms:modified></cp:coreProperties>'
  'xl/theme/theme1.xml':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office 主题"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'
  'xl/styles.xml':'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><color theme="1"/><name val="宋体"/><family val="2"/><charset val="134"/><scheme val="minor"/></font><font><sz val="9"/><name val="宋体"/><family val="2"/><charset val="134"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/></styleSheet>'