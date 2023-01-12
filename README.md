# protobi-msexcel-builder

A simple and fast library to create MS Office Excel(>2007) xlsx files(Compatible with the OpenOffice document format). 

Features:

* Support workbook and multi-worksheets.
* Custom column width and row height, cell range merge.
* Custom cell fill styles(such as background color).
* Custom cell border styles(such as thin,medium).
* Custom cell font styles(such as font-family,bold).
* Custom cell border styles and merge cells.
* Text rotation in cells.

## Getting Started

Install it in node.js:

```
npm install protobi-msexcel-builder
```

```javascript
var excelbuilder = require('protobi-msexcel-builder');
```


Then create a sample workbook with one sheet and some data.

```javascript
  // Create a new workbook file in current working-path
  var workbook = excelbuilder.createWorkbook('./', 'sample.xlsx')
  
  // Create a new worksheet with 10 columns and 12 rows
  var sheet1 = workbook.createSheet('sheet1', 10, 12);
  
  // Fill some data
  sheet1.set(1, 1, 'I am title');
  for (var i = 2; i < 5; i++)
    sheet1.set(i, 1, 'test'+i);
  
  // Save it
  workbook.save(function(err){
    if (err)
      throw err;
    else
      console.log('congratulations, your workbook created');
  });
```

or return a JSZip object that can be used to stream the contents (and even save it to disk):

```
workbook.generate(function(err, jszip) {
  if (err) throw err;
  else {
    jszip.generateAsync({type: "nodebuffer", compression: "DEFLATE"}).then(function(buffer) {
    require('fs').writeFile(workbook.fpath + '/' + workbook.fname, buffer, function (err) {});
  }
});
```

You can now provide the file path on save rather than in the constructor:

```js
   workbook.save("/tmp/workbook.xlsx", function(err) {
      if (err) throw err;
      console.log("open \"" + path + "\"");
   });
```

Further you can optionally compress the saved file:
```js
   workbook.save("/tmp/workbook.xlsx", {compressed: true}, function(err) {
      if (err) throw err;
      console.log("open \"" + path + "\"");
   });
```

## Use in browser

This depends on `xmlbuilder` and `jszip` which are included in the direct
```html
<!DOCTYPE html>
<html lang="en">
<head>
  <script type="text/javascript" src="./xmlbuilder.js"></script>
  <script type="text/javascript" src="./jszip.js"></script>
  <script type="text/javascript" src="../lib/msexcel-builder.js"></script>
</head>
<body>
<script>
    var workbook = excelbuilder.createWorkbook()

    // Create a new worksheet with 10 columns and 12 rows
    var sheet1 = workbook.createSheet('sheet1', 10, 12);

    for (var i = 1; i < 10; i++) {
      for (var j = 1; j < 12; j++) {
        sheet1.set(i, j, i * j)
      }
    }

    workbook.generate(function (err, jszip) {
      if (err) return callback(err);

      jszip.generateAsync({type: "blob", mimeType: 'application/vnd.ms-excel;'}).then(function (blob) {
        var filename = 'test.xlsx'
        if (navigator.msSaveBlob) { // IE 10+
          navigator.msSaveBlob(blob, filename);
        } 
        else {
          var link = document.createElement("a");
          if (link.download !== undefined) { // feature detection
            // Browsers that support HTML5 download attribute
            var url = URL.createObjectURL(blob);
            link.setAttribute("href", url);
            link.setAttribute("download", filename);
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
          }
        }
      })
    });
  }
</script>
</body>
</html>
```

## API

### createWorkbook(save_path, file_name)

Create a new workbook file.

* `save_path` - (String) The path to save workbook.
* `file_name` - (String) The file name of workbook.

Returns a `Workbook` Object.

Example: create a xlsx file saved to `C:\test.xlsx`

```javascript
var workbook = excelbuilder.createWorkbook('C:\','test.xlsx');
```

### Workbook.createSheet(sheet_name,column_count,row_count)

Create a new worksheet with specified columns and rows

* `sheet_name` - (String) worksheet name.
* `column_count` - (Number) sheet column count.
* `row_count` - (Number) sheet row count.

Returns a `Sheet` object

Notes: The sheet name must be unique within a same workbook.  No error checking or collision-resolution mechanisms are currently applied.

Also the sheet name is cleaned, replacing disallowed characters `[]\/*:?` with a dash `-`

Example: Create a new sheet named 'sheet1' with 5 columns and 8 rows

```javascript
var sheet1 = workbook.createSheet('sheet1', 5, 8);
```

### Workbook.save(callback)

Save current workbook.

* `callback` - (Function) Callback function to handle save result.

Example:

```javascript
workbook.save(function(err){
  console.log('workbook saved ' + (err?'failed':'ok'));
});
```

### Workbook.cancel()

Cancel to make current workbook,drop all data.

### Workbook.set(obj)

Represent and create an entire workbook via JSON data 

```js
    var three = {
      set: 3, font: {
        name: 'Verdana',
        sz: 32,
        color: "FF0022FF",
        bold: true,
        iter: true,
        underline: true
      },
      align: 'center',
      fill: {
        type: 'solid',
        fgColor: 'FFFF2200'
      }
    }

    var pojo = {
      "worksheets": [
        {
          "name": "sheet1",
          "cells": [
            ["A", "B", "C", "D", "E"],
            [1, 2, three, 4, 5],
            [6, 7, 8, 9, 10],
            [11, 12, 13, 14, 15],
          ]
        },
        {
          "name": "sheet2",
          "cells": [
            ["A", "B", "C", "D", "E"],
            [1, 2, 3, 4, 5],
            [6, 7, 8, 9, 10],
            [11, 12, 13, 14, 15],
          ]
        }
      ]
    }


  var workbook = excelbuilder.createWorkbook().set(pojo)
```

### Sheet.set(col, row, val)

Set the cell data.

* `col` - (Number) Cell column index(start with 1).
* `row` - (Number) Cell row index(start with 1).
* `val` - (String) Cell data.  May be a string or number.

No returns.

Example:

```javascript
sheet1.set(1,1,'Hello ');
sheet1.set(2,1,'world!');
```

Date values are recognized.  If `val` is an instance of `Date` then 
the data is converted to an Excel value (e.g. `new Date('2016-06-23')` becomes `42544`)
and a date format is applied in Excel. 

__hack__
For some reason, the generated workbook only applies Date formats when the fill is also set.
So when a date value is set, the default format is filled with a white background.
You can override that with an explicit call to `fill`:
```javascript
    sheet1.set(1, 4, new Date('04/01/2009'))
    
    sheet1.set(1, 5, {
      set: new Date('04/01/2009'),
      fill: { type: "solid", fgColor: "FFAA000"},
      numberFormat:"m/d/yy"
    } )
```





### Sheet.set(col, row, obj)
You can also set objects as shorthand.  If the properties match a method
then the method will be called with that argument, e.g.

```js
 sheet1.set(1, 1, {
      set: 'Red, bold, italic, underlined and centered with border',
      font: {
        name: 'Verdana',
        sz: 32,
        color: "FF0022FF",
        bold: true,
        iter: true,
        underline: true
      },
      align: 'center',
      fill: {
        type: 'solid',
        fgColor: 'FFFF2200'
      }
    });


    sheet1.set(2, 2, {
      set: Math.PI,
      fill: {
        type: 'solid',
        fgColor: 'FF0022FF'
      },
      numberFormat: '0.00%'
    }) // 10=>'0.00%'


    sheet1.set(3, 3, {
      set: '' + Math.PI,
      fill: {
        type: 'solid',
        fgColor: '99BB66'
      }
    })
```    

You can also set multiple cells in a Worksheet by passing a nested object or array of arrays
where the first layer is the columns, the next layer is the rows, and the third layer is the cells: 
```js
sheet.set({
      "2": {
        "4":{ set: "Cell B4", font: { bold: true, size: 14, color: "44FF22"}}
      }
    })
```

It's also possible to set data as an array of arrays.  
Just note that Javascript array indexed start at zero while Excel column/rows start at one.
So the first item in each array, at index zero, is ignored.

```js
sheet1.set([
    null,          // column 0 is ignored
    null,          // column 1 "A"
    null,          // column 2 "B"
    [null, "C1"]   // column 3 "C"
])
```
### Sheet.formula(col, row, str) 
Create a cell formula
```
sheet1.formula(3, 3, "A3*B2")
```


### Sheet.width(col, width)
### Sheet.height(row, height)

Set the column width or row height

Example:

```javascript
sheet1.width(1, 30);
sheet1.height(1, 20);
```

### Sheet.align(col, row, align)
### Sheet.valign(col, row, valign)
### Sheet.wrap(col, row, wrap)
### Sheet.rotate(col, row, angle)

Set cell text align style and wrap style

* `align` - (String) align style: 'center'/'left'/'right'
* `valign` - (String) vertical align style: 'center'/'top'/'bottom'
* `wrap` - (String) text wrap style:'true' / 'false'
* `rotate` - (String) Numeric angle for text rotation: '90'/'-90'

Example:

```javascript
sheet1.align(2, 1, 'center');
sheet1.valign(3, 3, 'top');
sheet1.wrap(1, 1, 'true');
sheet1.rotate(1, 1, 90);
```

### Sheet.font(col, row, font_style)
### Sheet.fill(col, row, fill_style)
### Sheet.border(col, row, border_style)

Set cell font style, fill style or border style

* `font_style` - (Object) font style options 
The options may contain:

  * `name` - (String) font name
  * `sz` - (String) font size
  * `family` - (String) font family
  * `scheme` - (String) font scheme
  * `bold` - (String) if bold: 'true'/'false'
  * `iter || italic` - (String) if italic: 'true'/'false'
  * `underline`: (String) if underlined: 'true'/'false'
  * `strike`: (String) if striked: 'true'/'false'
  * `outline`: (String) if outline: 'true'/'false'
  * `shadow`: (String) if underlined: 'true'/'false'
  * `color` - (String) font color as HEX RGB or ARGB value (e.g. `"2266AA"` or `"FF2266AA`"`)

* `fill_style` - (Object) fill style options
The options may contain:

  * `type` - (String) fill type: such as 'solid'
  * `fgColor` - (String) front color, as HEX RGB or ARGB value (e.g. `"2266AA"` or `"FF2266AA`"`)
  * `bgColor` - (String) background color

* `border_style` - (Object) border style options
The options may contain:
  * `left` - (String) | (Object)
  * `top` - (String)  | (Object)
  * `right` - (String)  | (Object)
  * `bottom` - (String)  | (Obtject)
  
If (String) it may be `'thin'| 'medium'|'thick'|'double'`.

If (Object) it may be:
  `{
      style: <string>,
      color: { rgb: <string> } | { theme: <int> }`

 
Example:

```javascript
sheet1.font(2, 1, {name:'黑体',sz:'24',family:'3',scheme:'-',bold:'true',iter:'true', color: 'FFAA00'});
sheet1.fill(3, 3, {type:'solid',fgColor:'2266aa',bgColor:'64'});
sheet1.border(1, 1, {
  left:'medium',
  top: {
    style: 'medium',
    color: {rgb: "FFAA8844"}
  },
  right:'thin',
  bottom: {
    style: 'medium',
    color: {
      theme: 5
    }
  }}
)

```

#### Sheet.numberFormat(col, row, numfmt)
Specify a number format by string or index.  You can now add custom number formats as well:



Example:
```javascript
sheet1.numberFormat(2,2, '0.00%');
sheet1.numberFormat(2,3, 10); // equivalent to above
sheet1.numberFormat(2,4, "$#,###.00"); // equivalent to above
sheet1.numberFormat(2,5, '" ABCDE "0.0%;" ABCDE "-0.0%;" ABCDE "—;@'); // not that you would but you could

```

The following number formats are built in, inherited from the original project.

```js
      0: 'General',
      1: '0',
      2: '0.00',
      3: '#,##0',
      4: '#,##0.00',
      9: '0%',
      10: '0.00%',
      11: '0.00E+00',
      12: '# ?/?',
      13: '# ??/??',
      14: 'm/d/yy',
      15: 'd-mmm-yy',
      16: 'd-mmm',
      17: 'mmm-yy',
      18: 'h:mm AM/PM',
      19: 'h:mm:ss AM/PM',
      20: 'h:mm',
      21: 'h:mm:ss',
      22: 'm/d/yy h:mm',
      37: '#,##0 ;(#,##0)',
      38: '#,##0 ;[Red](#,##0)',
      39: '#,##0.00;(#,##0.00)',
      40: '#,##0.00;[Red](#,##0.00)',
      45: 'mm:ss',
      46: '[h]:mm:ss',
      47: 'mmss.0',
      48: '##0.0E+0',
      49: '@',
      56: '"上午/下午 "hh"時"mm"分"ss"秒 "'
```


### Sheet.merge(from_cell, to_cell)

Merge some cell ranges

* `from_cell` / `to_cell` - (Object) cell position
The cell object contains:

  * `col` - (Number) cell column index(start with 1)
  * `row` - (Number) cell row index(start with 1) 

Example: Merge the first row as title from (1,1) to (5,1)

```javascript
sheet1.merge({col:1,row:1},{col:5,row:1});
```
### Sheet.autoFilter(filter_spec)
The argument may be a range (e.g. `"$A1:$J12"`) or `true` in which case the entire sheet domain is used as the range.

## Sheet.sheetViews(obj)

Optionally toggle grid lines and set zoom scale: 

    sheet1.sheetViews({
      showGridLines: "0",
      zoomScaleNormal: 50,
      zoomScale: 50
    })

### Sheet.split(ncols, nrows, state, activePane, topLeftCell)

Optionally freeze first rows and/or columns.  At a minimum specify the number of columns and rows.
The state defaults to "frozen", activePane to "bottomRight" and "topLeftCell" is calculated.

    sheet1.split(1, 2)
    sheet1.split(1, 2, "frozen", "bottomRight", "B2")

### Sheet.printBreakColumns([colIndexes])

Optionally set page breaks at specific columns
 
    sheet1.colBreaks([15,30,45])

### Sheet.printBreakRows([rowIndexes])

    sheet1.printBreakRows([15,30,45])

### Sheet.printRepeatRows(start, end)
Set rows to repeat on each printed page. 
Arguments may be specified individually or as an array of length 2.

     sheet.printRepeatRows(1,3)
     sheet.printRepeatRows([1,3])

### Sheet.printRepeatColumns(start, end)
Set columns to repeat on each printed page.
Arguments may be specified individually or as an array of length 2.
    sheet.printRepeatColumns(1,2)
    sheet.printRepeatColumns([1,2])

### Sheet.pageSetup(obj)

Optionally set paper size, orientation, and resolution:

    sheet1.pageSetup({
      paperSize: '9', 
      orientation: 'landscape' || 'portrait',
      horizontalDpi: '200', 
      verticalDpi: '200',
      pageOrder: 'overThenDown' ,
      scale: '50',
    })

### Sheet.pageMargins(obj)

Optionally set margins, units in inches:

    sheet.pageMargins({
      left: 0.25,
      right: 0.25,
      top: 0.5,
      bottom: 0.5,
      header: 0.3,
      footer: 0.3
    })


### Sheet.addImage(obj)

Add an image (currently only PNG supported, SVG in the works...)

    sheet.addImage({
      range: 'A1:B7',
      base64: 'iVBORw0KGgoAAAANSUhEUgAAAXIAAAFyCAYAAADoJFEJAAAAAXNSR0IArs4c6QAAAHhlWElmTU0AKgAAAAgABAEaAAUAAAABAAAAPgEbAAUAAAABAAAARgEoAAMAAAABAAIAAIdpAAQAAAABAAAATgAAAAAAAADcAAAAAQAAANwAAAABAAOgAQADAAAAAQABAACgAgAEAAAAAQAAAXKgAwAEAAAAAQAAAXIAAAAAL2csFAAAAAlwSFlzAAAh1QAAIdUBBJy0nQAAQABJREFUeAHtnQl8HMWV/6u6Z0aSjU9sy5IMmCMBc2zCcuUAsrAEbNlaQg6HJCzB5DAL8ckNgWizId4NxBeYxIT8bZM/bNYJIcG2MCFghyVLEnACDpgNxBzGkiwbfFuWpo/aVz1qaSSNNDNSX9Xz6w9GM909dXyr9NOb169eMYYDBEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABBwCHBxAQBkCq1fr1RVbyhg7YoiW1EbaBh8ukvpQYYhhmsZGCiaGMa5VMGYPYZxVcMHLhWA6Z0JnjNM/Osu5Te8t+p/FhW0yoR2ms63MZoeZLlq50PYJYe/jmjjI0/yQlTD3taWG7T3SaG/7W+3sdmVYoaElRQBCXlLDHfHObqhP1OwZM8JO8XFMN44mya1mXBzFBD+KxHY8tX4safGRjNvD6ZwU6zIS5iTXpU7LvtH/3J8db+WPfg8hr9L/nJ8dr236k2Ba8o0U7ozQC7aX7nmfc7ZTcN7IBdsmuGgUFm8UuvVuWui7d9fO3i8/hAMEgiYAIQ+aOOpjjCzrCRXbxws9cZTF7A/qTDuRdPRkEuGjhC1quMZHkkCX82SiixaZ1oz+0XXnp3ztHO7Prju9eeX+Qej4SW2iPxL0T/7M/LXItCdtyPoOksjvYhprFDZ7i+7YQlb96zYXbyR11rh98vzd3jQKpYBAbgJyVuIAAf8I1Ndr1R8dVsNE8kRu2x+yOf8Q+TdOoQqPoX+jSay5I45SoG27t1D71zJvSpbiLg8SeE7+Hacv8hxZ9LZpmpxrzUKIN+mPwMt0x2YS+peTQt+6bdq1ezIfxP9BYPAEIOSDZ4gSsghM3LCi/LB58AO6ZZ9J7oePkD6fTrp2PNO00TwhLWwSbMsVbNt5m/XxeL3ssOAdgddJ5KXLJm1Qp1kTCfsWEvZNwrJ+L7j5cvMLB99l9fXyGg4QKJoAhLxoZPhANoFTVten3h8+5gNkbX+M3B3nk3adQZ6H43gyKf3XGSvbIn+ztLhxdLpnuBR2suAdX7xp7SVWrwnOnueWeFazrRe3181vBC4QKJQAhLxQUrivk0Dlk3eP01nZOWRcX0wCfS6p00k8lSh3XMfS2qZ/nT7szk/hRZ8EpFtGPrAlcXeE3bLep3tfoucBGzVNPG3uMje3XHnjoT4/jwslTwBCXvJToDAA4xqWHk86cyGF9E0hlf4oWZTjmRQfsrYh3IUxLPgustQ7Lfb2tPwq8zr52TeS3D/BbON3TXU3vFdwWbixJAhAyEtimAfWyaq1iyaxROJiirueRhb2OTyVHCZd3I7VSA8mcQRAgNxTjqi7fzRNq5G++Wy0TLGGHpxu3Hnp3JYAWoEqIk4AQh7xAQq6eZVP3nusbrMpZGZ/hsIwPkLiPYRCSZgwyM/tV6hf0J1UuT76WsQT9E1IPnOwrCYh+FM0Lo8eTrY+u+eTt+xTuWto+8AJQMgHzi42nxz11PIRQ630RRRT8UXS6gvI3z1KijbEO+JDLF0wMtZeLl4S9lbbttdoXKxuPFDzRzZ9urOiKeI9QPM8IgAh9wikisVUrVt4pqYnvkDG3WWarh8rY6CFaSLCRMXBJNeLtNQpvNEi18sfqAv/qRnmY4h+UXEwi28zhLx4Zkp/4ui1C0YZ2tBacr1eRbJ9PkulUh2REkr3C43vICB96tL1IkMb29O02pT/ilLNrGr6w67/QZx6fGcJhDy+Y9utZ2PXLzwhaelX0IOzL1GY2wkyxlsYZH3D792NU6zedLheRJp8L8L+nZbQftRqiseREyZWo+x0BkIevzHt1qOahqUfIW/318lC+xT5U0fB+u6GpzTeuFY69ZbG/3V6u0pw4ydNl1z/bmkAiH8vIeRxHGNB+U0aRv4jWd2zKNnTZJZMJoVByZ2wujKOo11cn8jtIhcfkS99Jy3keohCXx5smjL3r8UVgrujRgBCHrURGUx76usT1WcfOY1zMZeWe39C5jaB+2QwQGP8Wel2SdH8aDdk6t3VJrPv3Tll7uYY9zjWXYOQx2F4KS1s1ZAdl9LWCXPJhXIel1spOP7vOHQOffCVgEwPkExKC/0QPRj9KU+n722su/5lX+tE4Z4TgJB7jjTQAnnN+qW15EK5UTD+CZkz2xHwQJuAymJBQPrRUyToBgk6Y4+QL31R87R5r8WibyXQCQi5ooNcte6+87lm3UqrLyfzBIWapSkCBQcIDJaAK+hpcx9FND0oRPre5qk3vDPYYvF5fwlAyP3l63nplWuXnKrr/Baywj9P8cIJ5yEmMsR6zrnkC+z0oZs7BLOXlpXbP3j7gnl7S55LRAFAyCM6MD2bNebRRVWpCm0+bVLwNZZMjCCfJmLAe0LCe+8JOLldnIfmr5GgL2jeX/UIlv97j3mwJULIB0vQ788vX56sOurwVSTgt5EPc6Ij4Agj9Js6yu9BwFktSm4XYdtP0abY32qc8o3ne9yCtyESgJCHCD9f1dVPLDmXHkF9hyX0T8gYcNoWLN9HcB0EfCXgPBC1rDbpP9fb+YJ3L53V5GuFKLwgAhDygjAFe1PlL2gHnork7bRx7zWUDzzlWOHBNgG1gUDfBNwHoob1Nhf2vzb+YfdDyOPSN64grkDIg6BcRB1khV9OK+6+w8uSx9NiDfjBi2CHW4Ml4GxPJ0NeLXutJqzbt2NBUbADkFUbhDwLRpgvKx+nDR1SYgFZ4Z+nNKSZXXjCbBDqBoECCZDRIXOi7yVBv6vpYNVSehiaLvCjuM0jAhByj0AOuBhayVPdsOQqltD+jR4o1WSiUQZcGj4IAuEQcDItUvrctLXRttpvaJ52w6ZwGlKatULIQxz36qeWHM0tfjdtYjyddkx3NjIOsTmoGgQGTcB5GGqaB2hv139r2la2mM2cSf5BHH4TgJD7TbiP8sc3LP0sGTHfpzwXR9MGAH3chdMgoCABaZ1TlkXbNNezVmt+86ex1N/vUYSQ+024R/kTNywamW5L3MU09i+0OpM7+y32uAdvQSAOBKTvnHK27OSWdWNj7VxKmYvDLwIQcr/I5ih3/BOLz9K49gP6+nkGIlJyAMKp+BGQK0PJXrGFeEDfZ9y6ffr83fHrZPg9gpAHNAbV65dcRwmuvktbrQ1HgquAoKOaaBCQcefSOk+bL2mWdc32qXPl5tA4PCQAIfcQZq6iJqxeONoert/D9cQMCs+iXXroHw4QKEECtNUgxdXae+20fXPztDkPlCAC37oMIfcNLWPVDUtPZzp/gHbqORMPNH0EjaLVIUAPQmkDC7mIaHm7YDdhI2hvhg5C7g3HXqXUrFs4nemJZbTEfgyW2PfCgxOlTIBUh5el5K5E/20YbV/ZVXfTG6WMw4u+Q8i9oJhdBu2bOf6cUbfTQ8076CmPzpDoKpsOXoNAJ4FMAi5zG7PY15pqZ/+68wJeFE0AQl40sr4/MPExCi0sp6iUZPJybHrcNydcAQGXgIw3F4K1MtO4qWnqvGXuefwsjgCEvDhefd497rHvHZ+oKH+IrIyPiTYs8OkTFC6AQE8CcgGR9Jub1j1V75bftgmrQXsSyvseQp4XUf4bqtcu+RjlSnmIrIvj4Q/Pzwt3gEAvAm5q3LTxM42b12yfjHjzXoz6OQEh7wdOIZeq1iy8jKzwB2kHn9HYwb4QYrgHBPom4DwENYznhJW+Aps+982p5xUIeU8iRbyvWrd4ppZILKZ0V+V4qFkEONwKAv0QyIi5uYVynH8BOc77AZV1iYI6cQyEAIn4Nyk+/AdCCIj4QADiMyDQBwG55oIntJOFlmig37Pz+7gNp7MIwCLPglHQy/p6bfw5o/9dT+g30sMZ7OBTEDTcBALFE+BJimixxW5aCfrPO+rmNBRfQul8AkJexFifsnp1avcRzfdqZcmvO/lSKG4KBwiAgH8EOsITDzLD+lrTtDk/9a8mtUuGkBc4fhM3rCg32vb/iKVSV4g0hRdCwwskh9tAYJAEZAZFxtsp1vyaxqnzVg6ytFh+HEJewLBWr1k+hCXafsxTqcuRM6UAYLgFBLwmIGPNOUvbaesbzXVzf+R18aqXByHPM4KVT949VLNSq7Sy1Gcg4nlg4TII+ElAJtyirViEac9qnjrnh35WpVrZiFrpZ8QoBW2FbpetgIj3AwmXQCAoAjIFtGAJimi5r2rtkq8HVa0K9UDI+xiliSvqy+1hCXKnJD8HS7wPSDgNAkETyIi5rqX0+6rXLZkRdPVRrQ9CnmtkXlyebK8cuZxE/AsQ8VyAcA4EQiRAYk5hiUmyzH9Y1bD0SyG2JDJVQ8h7DoUQvLqlbZFWVnYlRLwnHLwHgYgQkGLOWIq2Tnygeu3if4pIq0JrBoS8B/qa9Uvvov0Fr3NCDHtcw1sQAIEIEZBbJzI2hKf0VZUN914YoZYF3hQIeRby6oZFN9GOPrdmcolnXcBLEACBaBJwNm7hI3WdPVK19p4zotlI/1uF8MMOxuMbFl9JCbB+TJsjJ5iN1T7+Tz3UAALeEXA2djatrYZgl+ysnb3Vu5LVKAlCTuNUtW7JZHpw8nNyug3FLvdqTFy0EgR6EqDgBMYM8wVLsKk7amfv6nk9zu9L3rVSvWbR39MDk4do2RhEPM4zHX2LPQFnU5dU8iydi5VyDUjsO5zVwZIW8ppH/2MCS+gPkzU+lslMhjhAAASUJuBEmqVStfYR+veV7kiRjS9ZIR+7etkRoqJsFX0dO8nJZFgkONwOAiAQTQLSMief+b/QFow3RLOF3reqVIWcJ4aai3h52YWIFfd+UqFEEAiVAKWXFhSaSGGJ35VbMYbaloAqL8mHnbTryDwtlVyIPTYDmmWoBgTCIEAxibQEdJcp7It2Tpm7OYwmBFVnyQm5jFDREvpjQtjlCDMMapqhHhAIh4CMZCFXy0vMMj7ZVHfDe+G0wv9aS8q1MvZXC0+g3e4fFJw2S0asuP+zCzWAQMgEHH95eerDTEsuY6tXk4kez6NkhFxuDpFI6A/SPoA1iFCJ52RGr0AgFwHRRps5lyWnVw9tmp/rehzOlYyQC+3wd7Ty1CecWNM4jBz6AAIgUDABYZqM0m98u3rtoosK/pBCN5aEj7z6iSWXM01/mFZtagwbJis0PdFUEPCOgLORs2W/pZnmedvr5jd6V3L4JcXWZ+SinfDrJR8QjK+mlZvDsPzepYKfIFCCBOi5GC9LjaJc5sce+MBHHmUbNzrpE+NAItaulYkrVpTbJv8hTyQqmZMlLQ5Dhj6AAAgMlIBcN8KTycuqzho5a6BlRPFzsRbydOW+W+khx4Xwi0dx6qFNIBAOAUFGHeVX+rfx65eeHU4LvK81tj5yWvRzPqWlfVLYFC8Ov7j3MwclgoDCBHgqwWhB4KZ0uu2C9y69+YDCXXGaHkuLfMTaBaPIJ76MaRwirvoMRftBwAcCMr8ST6XOSCXL7/Sh+MCLjKWQD9WGfptCDU/FEvzA5xMqBAFlCDiLhXR9TvWapRcr0+g+Gho710rHJhFraOUmfXfCTj99jDtOgwAIEAG5sxBZ568lROvHt027dY+qUGJlkR+99v5RnLOFtAwfIq7qjES7QSBAAvJbOy9PTbK0CqVdLLEScou3f4viRCfBpRLgbwKqAgHFCThRbbp+7fgnFv2Dql2JjWulau3C8yhK5TfkTUnBpaLqdES7QSAcAo6LxTD/bOmjzmu55MpD4bRi4LXGwiKX+/NxTb+b6TpEfOBzAZ8EgZIl4LhYylKna9aeuSpCiIWQ0/5811He4XPgUlFxCqLNIBANAlI/NE27sXLtklOj0aLCW6G8kE9Yv/AExtmtcrUWDhAAARAYMAGbUq8k9RE6FwuYqFdKG5VqbK4Bsi3tXymwfzSjPfpwgAAIgMBgCIh2g8Q8ObWqYaRSe30qLeTVDQsvZsnEdORSGczUxWdBAASyCXAmKIqZf2fU6n8fkX0+yq+VFXL5gJMJ7S7OOWLGozzD0DYQUIyAMCmpVnnqpPJhQ5TJkKiskFtDtSspZvxMPOBU7LcEzQUBBQg4USxMzKt8/N5jFWguU1LIq9fcM4Ys8VuxUYQKUwxtBAEFCdCDT/nsTdft21RovZJCLvTkXLLGj5FfgXCAAAiAgB8EhEEPPnX+z+PXRD9vuXJCXrl20XH0IOJauFT8mLooEwRAoJOA3BoumSjTdPFNWi0e6VXwygm5rmvzadefUXCrdE43vAABEPCJQEceltqa9Usv9KkKT4pVSsjHNyw9mTaM+LLzlceT7qMQECg9AjK5s40Uz4UNPMGibeF0YbPb2Yb6RGEfCv4upYRcY+Im2jj1CJqFwZNCjSAQAwKCkbuA+jEsWcZMUicc+Qk4G1Ak9AuqDw+fkv/ucO5QRsgr1y08jaxxLP4JZ56g1hgQkOZPmqIxZh1/Jlt91mXs6Irh9B4BAwUNrU5//oR+Y1StcmWEXGfkG08laREQrPGCJh5uAoEsAhkRt9jc489iN3/wY+zDIyvZyjPq2DFDRrB2iHkWqdwv5R6ftIr83JrWkZNz3xHuWSWEvPKXZI1rOqzxcOcKaleUgHSnSMtbivjtJ33cca3IrpwyfKwj5hMh5gWNLNcYJ2dUJK1yJYRcT2izeCoxBNZ4QfMNN4FAJwEp4u3kTpEi/s2TziUR7x5FdyrEvJNVvhfSKqfNa86rajvyH/PdG/T1yAv5uIalx9NT4+mIVAl6aqA+1Qm4PvGMJX5un93JFnP4zPvElLmga1yzrTlRiyuPvJAnhD2TlSVHIFIlzwTDZRDIIuD6xOd0WuJZF3O8lGK+Cj7zHGS6n3IWIib0i2rWL/tI9yvhvou0kI9dt2w8Rapc6TxoCJcTagcBZQi4Ip5xp3T5xPN1AD7zfITougy2SOhJYZuRyowYaSFPMfOLlFOlEqs4C5hguAUEpM44PnGLdVni3X3i+SBlu1kQzZKblmNYalrdmHXf/2DuO4I/G1khr15TP4SyG3yFYQu34GcFalSSQLZPXD7YHOjhullkNAt85jkoklVOaUKOSHH96hxXQzkVWSEX2sgpPJE4GRkOQ5kXqFQxAq47pcsSH1wHpJsFPvO+GXYk7btCptTu+67grkRTyCnTGB1fZ1pxXwuDw4aaQCA6BLLjxL+ZFSc+2Ba6Yo448xwkaY9gyldeIxLJT+e4GvipSAp5TcPiD5GQfwKpagOfD6hQMQJunHiXJe6t8QMx72dCSBeLzWawDRtCT6YVSSG3mXYlS1FWHyzH72cW4VKpE/DKJ56PY7aYw2feRUuYtGxf186qOvTnj3adDedV5IS85hf3HUlbWH/OgRQOE9QKApEn4LVPPF+HpZgjN0sPSjQItPGEzjX9qh5XAn8bOSG3U+YUSo41gZEPCgcIgEBvAt1FvPA48d4lFXcGoYm9eXUYnHVj131vfO+rwZ2JlpCTy4mOLwXXfdQEAmoRcB9sSp/4HTlyp/jdm+zQRMSZE2350LMsOTZhp+r8Zt9f+Xp/F4O+VnXWokmUV+UuWgCUDLpu1AcCUSfg+sS7HmyG0+JxZUPZx4+cwDa+9w7bnT5M+xNHyx4MmgrXKC2iYMMOfOCcn7CNG+UwBX5EagSEpn2W/rohy2Hg0wAVRp1Ad3eKzGIY7gGfeRd/udaFFi9+dNzpI0/tOhvsq+gI+YvLk9SYy5gJ33iwUwC1RZ1AdxEPzieejwt85h2EZBgiRdklkuxT+Zj5dT0yQl65s/0MpmmnYSWnX0ONclUkELZPPB8ziHkHIZlKRPBPn9CwtCwfMz+uR0bIaUu8T1MoDwXWh+Ji8oMtygSBQRHobokPPHfKoBpRwIezH4CWapy5kEKua6fut8UZBSDz/JZICLlMkEX5xuuQIMvz8UWBihLoEvGzO3b2iXZH3EVDJbsHKA0Y7WKmU1R5KO6VSAg5S4z6e8qr8kHnr1q05ytaBwK+E+gScbk9W3R84vk67j4ALdncLPKhpy2msBUryvOx8vp6NIRcsDp6WKDBq+L18KI81QhE3Seej6d0szx0Zh078YjRzBKlFbggDVGu80mVlQc+nI+T19dDF/KOhwNT4FbxemhRnmoEulvi0fWJ5+M6adhYNmnYGGaWWq4kx72S1MkinZaPkdfXQxfyQ8I8jYJiJwksyfd6bFGeQgS6RFwNn3hfaA3bZtdv/g37ZfPrrEyL1HrDvprs7XlHx8QlbEN9oBkRQxdyzvSLKK9vApkOvZ1PKE0dAl0irpZPvCdhKeI3vfIMW7ltM0uVoogTEBk+rXHt1PGHRwe6DVy4Qk4bSDCNXYw9OXv+SuB9qRBQ3SfujlNGxJ9mq0jEK/RE6CtP3XYF/lO6k1KJco3ZFwZZd6hCPv6Xi4+hfp+BaJUghxx1RYVAxhK3OzdKjkq7im2HFPGbX32GRPwvjogX+/nY3S8HVvDJQfYrVCHXUuyjWjIxnGLIg+wz6gKB0AlkRDx7t/vQmzSgBrgivvKdjCU+oEJi9iFpmAouzhrfsHRsUF0LVcgZ0y7EvpxBDTXqiQqB7iKuTpx4T37ZIl5O7hQcHQTogaem6+NIXE8PikloQj5xRX055X78OKJVghpq1BMFAnHzibuWeNjZGKMwtt3akJTxG/YF3c75+CY0IW8dM/pEzrUTsBOQj6OLoiNFoLslrm6cuGFb8Innm1lywwnO/4Gt/lwgMZihCbmu22fT090kwg7zzQhcjwOBjIib9GBT/Tjxm1/dwFxLPA5j40cfHD+5YCdXV5xb40f5PcsMTcjpacB5lGamZ3vwHgRiR6C7JQ6feOwGOFeHKIBDS+rDWUIEslw/FCGfsHphBefsTOQezzUDcC5OBLr7xM8j00VN48WNE3ctcTV7EfDMSpBXxWaB+NBCEXKznB1HLpXjsRAo4ImF6gIl0N0SD+T32Zf+OT5xWrGJOPEi8ZKfXHB+Dquv911nfa8gV9cpNOdDlO0wBf94Ljo4FwcCXSIeE594x4rNOIxNUH2QEXn0zWVS5d8NHeN3naEIuQyWp23d/O4bygeBUAh0ibj6uVPkik3pTkGc+ACmEq14JZ0bo5UlTxrAp4v6SPBqSl8zuOBnwK1S1DjhZkUIZPvEv3mS3O1eTW8yfOLeTDjavpImAPd9YVDgy7HGnz36SEJ0IhYCeTNRUEp0CHRZ4qq7U2Sc+Ab4xL2aWpyd5VVRfZUTuEUuhHYcGSlHwj/e15DgvIoE4iPiMgEW4sQ9m4Ny1yDGTvY7P3ngQq4z62S5SSmE3LOpgoJCJuCK+OzjpU9culPUPJA7xftxoz08ZaETjzw4vNL70rtKDFzIaZaTv0jVqd4FDq9AQBLI9onfobyIP925YhO/oR7N78wDz5FJnZ3gUYk5iwlWyIWj4CfjQWfOscBJxQi4lvgc5S3xTO6Ule8gn7gfU5Anda7Z2il+lO2WGaiQV/76oSG0kcRxQv6VwgECChNwRXy2kztF9WX38In7OhVpGTtl0DrZzzoCFXI7vaeaczEOG0n4OaQo228CXSJ+Fsu4U9R0RMAn7vdM6Sg/Y7j6uodnoEKe1MRErutD8aAzoAmEajwnECef+E2vwifu+QTJUaATas3FsdVrlg/JcdmTU4EKOX2/+CDTA0nP6wkcFAIC2QSyLXG1o1M68onDJ549vP69lhsyC1Zp8dbxflUSqJDT0vwTsLWbX0OJcv0k4Iq4zCd+x0nwifvJOnZlyxBEXRumMy0eQk4DNBFuldhN09h3qEvE3ThxlX3iGXcKcqcEO225XDqj2cf6VWtwS/RpyyPKO3EMtnbzayhRrh8EMj5xm3b2cUXcj1r8L9PJnUI+8VVwp/gPO1cNlCRQ49w3IQ/MtTJh+EdHUP8qhfQX4QABBQj0tsQVaHSOJnbusQkRz0EnoFOkezL02q/aAhNy6sBo+jcCoYd+DSXK9ZJAl4hTAqwT1feJr6JUtBV6cF/AvRyLWJTlPPDkR/nVl8CEXNiJMRQUj9BDv0YS5XpGICPicqPkDneKXNCh4JEdJ14GEQ93BOUDT84q2fLlST8aEpiQc2HXcD2w6vxghTJLgIAbJx6HBFg3dWwKIS1xNf8UxWfCdSTPGl1zvDHcj14Fpqw259WIIfdjCFGmVwRcd4oUcbUTYLlx4nCneDU3Bl1O5tngKLPNkM8KPT8CE3J6yDmOXCuedwAFgoAXBLqJuOMTV3OuZtwptCkEfOJeTAuvyyin2JVRXhcqywvu6YfMsZL5q+RHP1AmCAyYgOsT77LEFRVxYbNbsMfmgOeBrx+U2qdxjYxZXzZiDswi50yHkPs6U7oXnrYtZtEvNo7+CWT7xO846TzyJSsq4pSY6ea/PMNWdFjiavai/7FS/apcFETGrNpCzpg9FhZ5MFOxnUT8ksrj2LFDRzEp6DhyE+jmTlF6UwgrY4lvw273uUc6ImdpUZDaQr56tcyUNQxC7v+EOmyZ7LM1J7EVf1/HfnR6LTu6YjjEPAd2KeLyD57jTomBT3wlLPEcoxy9U/RNyRcfeSCulbFsVwUFUVZQBjAcPhI4bBls+oRJbPFpn2RJ+ut/6vCxbMUZdewoEnMpWjgyBDKWuIwTPzMTnaLoQ3gjyyeOOHEFZjdNPMHZSD9aGoiQJysOl1Pjh2J5vh9DmClTWuKfn3CyI+LZCZH+bsQ4tpLEHJZ5hlNsfOIk4vCJ+/f75E/Jzl6X6go511NSyMkqh0nuxwSRIi4t8UVkiWeLuFuXFPNVEHNn9slnBrMdS1w+2FTzcHKnvPIMW0k+cSy7V2gMpfwJhS1y07ZS1IUy6Lj3k84VcelOySXibo2nlbhlLn+HOn3iij/YvPlV7LHpzmvlfnLhyy5BgbhWdJGQQu5LjgHlBtLDBkuf+OdrMj7x/kTcrVKKeSn6zKWIS0u80yeuqC2e8YlnRLyQ8XbHHT8jQoBiyem/Mj9aE4iQc80k1woFw+PwjIDrE1/0d/1b4j0rLDU3S6dP/Di57F7hOHH4xHtOZSXf0zqFMlZf77kWel5gTrpcLycjiOqCjzwnnyJPuu6Uvnzi+YqTlrn0mctoljjHmbuW+OzjKDpl0rmK2uGMwSeeb0Yrcl2u7uSs7JSTvV9RH4iQ2xq5VQQF3uAYNIGMJV64O6WvCl0xl9EscQxNlCLu+MQdS1xxEYdPvK9prNx50vLEobHSqPX28LzAXM0TwtKRMCsXmeLOOXHi5BMfqCXes7bsB6BxEvOMiHfEiUtLXOk4cfjEe85bld/TVNTMXcM9N2oDEXLaPVqu7IRnZRAz0PWJLy7SJ56vStcyPyYmbhbXJy43hVDZJ25KnziFGCJ3Sr4ZrNB1aWFwrhkjLM911/MCc2ElrwpZ5Lmu4FwhBAbrE89Xh2uZq+4zl78nsYoTp2X3iE7JN3vVuS7nJ211qdv7hniuhoEIuTqoo9dSr3zi+XrmWuaq+swz7hQ3d4r6PnFY4vlmLK5nEwhEyDkXFgJWsrEX9tqNE/fKJ56vVinmMs5cNTHPiDj5xGV0ilzsA594vqHG9RAIOGa4xqxkRcLz/NKBCLnFOjI2ef6FIoTRCKhK1ye+0GOfeL7mu3HmqvjMu/nEJ6kbJw6feL6ZGYPrUv8Es7XDhxwvi5c9CkTIOdfJIve87V5yiFRZjiXekTsljFwars886om25IyKh08882BTpqKFTzxSv4qeN4Zk0E4c2u+5GAYi5JrNDPq+63njPaccgQJdSzwod0pfXc4W8yiGJnaJOEWnnKi6TxzRKX3Nw7idJ6PcHDqUqelaYcJqk18p5LImHH0TyLbEo2CZRVXMpYi32yZzVmyq7BOX27N1LPaJwnj3PTNxxRMC8tkNZ22vbmGmJ+VlFRKIRS60ssPkHPL8r1BWP5R/mbHET/FssY9XQKSYy+X8UfGZx8on3rFRsnSfwcTxasZGuxzakyFNuVY818JAhNyy2g3CK//hyEHAWbHp+MQviqSPNNsyDzM3S5c7RUanqJxP3GY3yXzi8Inn+G2I8SmyyOm/dj96GIiQJzQ9TY1vh9nRewhdn3i+fOK9PxnsmWwxD8NnnhFxcqc4KzbV9onf0iHisMSDncMRqe2QH+0IRMiFlW6jxpN7BV8gswcxaj7x7Lbleh2WmHf5xOWyeynias4jN5/4Cux2n2t6xf+cM235Pj86GoiQG4crpJAfUnWhhh/gpVV7eU30fOL5+ur4zM8Mzmfu+sQdSxxx4vmGB9cjTYDLdZF7/WhiIEK+i40la5wscjUNKc+5Sz/zJ8cdy+798CWR9Inn6/BpwzMbOvudmyXbJ36n0tuzIU4835wqieukfxSEra6Qs+nTLYpaOSA9/TgY04nDGwd3sz/v26EsDscy93E5v+sTn9XpE1dz7shNIW6h6BTkTlF2qnvacJrXezwtsKOwQCzyjrp2QcgzJHTa9W5b63521Ytr2At7mvwY10DKlGIuc7PI0EQvH4BKEc/EiZ/FMpa4oiJOEbe3UJy4FHHEiQcyJaNdCa0bIA18z49GBibkgmstEPKuIUxqGmtpP8RmbFrL/qiwmMvcLCvJZ+7Vcn74xLvmCF7Fi4AwyTGhupAz24ZF3mNepmi/jZb2Vna14mIufeZe7AEKn3iPCYK38SEg3coWfUWzbLUtci74TiTO6j0vUx2WuRRz1d0srpgPxM3SKeKde2wq6k6BT7z3JMeZDAFanm8zW20fudB5I/01wpDmIOBa5tLNorqYrxyAz7zLJ+7udq+qiMMnnmN645QkkAn02JMoT6obRy77oQnWJCzyEeHIScC1zEvNZ+76xJ09NlWPE0d0Ss65jZOk45pjnOxu3Lp5vx88AnvYyTXzPXKtHMIDz76H0bXMv1IiPvNOd8rxHTv79I0m0lcMmcUQuVMiPUahN04KuWAtbOYDvuScCkzICeRu+rePZf4yhc41qg2QlvkOimbJ+Mybo9rMvO1y48zloqFcPnPXnTILPvG8LHFDDAhI1wpn2/zqSWBCvn3/89I31IJl+vmH0rXMr96kfpy59Jn33APUFfHZJOJ3TlI4dwpZ4ogTzz+fcQcRICGnvXXe8otFYELOpv/MIn/oO0wPrkq/oAVRrmuZx8Fnnp3P3HWnSJ/4nYr7xLFiM4jfhJjUQX/0hc3f9Ks3waoq52/BR174ULqWufSZxyGaRbpZDplpSkUbD584VmwWPpdL/U5nMZBgb/vFIVAh5zbbymxpk+EolEC2Zf7CHrV95j86fSqrn3S+2qloHXcKcqcUOn9xHxGQzwUt+4CV1Hz7BQ5UyOmx7esMIYhFz21pme90HoCq7TM/Y9R4Nu+Es9XNJ54l4sidUvQ0Lt0PZB50tujpVItfEAIVcsPmbwvLRgjiAEYzSWLuRrOonJtlAF2PxEdMWl19M+LEIzEWqjWCy+eCgr/VVDez1a+2ByrkWmpUkxBsJ0IQBzac0jLfQblZMnHmvn1LG1jjYvwpxInHeHCD6BqFFNPxup9VBSrkLRdf2UrfMt7kmY752a/Ylu36zDOhiRBzvwdaijiiU/ymHPPy5WNBIbb42ctAhZyco4LZbAuDkA9qTF2f+QzF48wHBSGAD2eLOHziAQCPaRXCMIVIslf87F6wQi57wtmf5VpVHIMjIH3mMp+5kwJ3t7qbUwyOgn+fhk/cP7YlVbI0Wi17bzptb/Wz34ELucX0LSJNGdblk1wcgyLg+syv/pPcnAJulkHBzPowfOJZMPByUAScZFmcvf3+Eft9i1iRDQxcyPkh+01a4vQ+hHxQ86Pzw9JnnrHMZWgixLwTzABfZLtTKvQEfYHEAQKDIKDr0v+whV1Qbw6ilLwfDVzId2zZ/T6F4vzVCcnJ2zzcUAgBaZl3iTncLIUwy3WPmfVgEz7xXIRwbkAEBHthQJ8r4kOBCzmrr7cFF5vwwLOIUSrg1oybhXzm5GZReTl/AV315Rb4xH3BWvKFkhuZUkzxP/kNInghpx7Rtm8v0B6efvet5Mp3xLyN9gB1xBxulkIngOsTR+6UQonhvoIIyAedtv2enW7/a0H3D+KmUITctqyXRdpIw08+iJHr46NOnHnbITbjT/CZ94Go2+mMT3wDkyIOn3g3NHgzSALSfSw4e61l8yFfNlzObl4oQp5oY2+SiG+FeyV7KLx7LS3znSTmVztiDp95X2SlT/xWZ9n9yww+8b4o4fyACZCQ08Py30t38oDLKPCDoQj59unzDwvGX+AJvcBm4rZiCTi5WaSYOylw4Wbpyc/1if8/WOI90eC9VwQoypqyvf7Oq+L6KycUIc80yH4OC4P6G5rBX8s8AG2Fm6UHSsedQntswifeAwzeekeAUtfahrWffCsveVdo3yWFJuQWT/2BpU0DfvK+B8eLK06cOXzmnShdnzgs8U4keOEDAU7x47QY6NWmw881+lB8ryJDE/IhO3a+Lmz7b9j6rdeYeH6i02de4rlZpIjDJ+759EKBuQjIB522/Vu5xWWuy16fC03I355R30aBiL/DwiCvhzR3eaWez1z6xOVGybDEc88PnPWYgGGSs0Hb4HGpfRYXmpBnWmQ/g63f+hwbzy+4PvOr/7SupJbzZ3ziMsQQ0SmeTyoU2JsAWeO2abdQqAolCAzmCFXIBUv+j22Y+7HRRDCDLWvJ+MwPlswD0Iw7RVriLyNOPLhpVtI1Of5xzl7YUTt7V1AgQhXy5trrtlESxE2y4ziCI9DdZx7f0MRsEUeceHDzq+RrcjKtiSeD5BCqkJMTSW408WssDApyyDN1uT7zGTHNzdLlE4clHvzsKuEaZXrutNlmM+2ZICmEK+TUU8Gs34h0Wj4ZCLLfqIsISMu8xVkBGi+fOXzimN5hEZCLHG1hv7KjYreve3T27F/oQj6UJ/5CVvlriF7pOTTBvM/kZjnoLOd/ca/6bpaMO+UZ+MSDmT6opScBZxtL/qTf+cd7VdvzRNDv/1Y7u53qfILBTx40+s76nGgWssy/StEsr+4P7PlMZ/1evejyiW9G7hSvoKKcwgmQU4GSAVo242sL/5A3d4ZukTvd0NgaAkD996ZTKKV4AlLMGw8fYF+mRUOvKCjmGZ84LPHiRx6f8IqADNqgRUCvtbQMC2RZfna7oyHk5p4/kWPpdUSvZA9N8K+lmL/duo9dRWKukmXe5ROHJR78rEGNnQTIP0626BNsxgxa7BjsEQkhb6qrb2WaWAP3SrCDn6u2MhLzd0jMpWWugph3uVMyi33wpS7XqOKc7wQct4ppmYL/0ve6clQQCSGX7bKE9gthmLRBKX4Vc4xToKdcyzzqYp7JJ55Z7CPjxDFzAp0mqCyLgONNsOxXxh5b9WLW6cBeRkbIW8aVbaJtkf6CHOWBjX2/FUnLXLpZouoz78onjjjxfgcSF4MhIIM1uPjFq6dMTwdTYfdaIiPk7MyZ9LSTPcYS0WlSd1Sl984V86j5zOETL725GOke0xoYCtZot4T+WFjtjJRqctv+uWg3WrE4KKzp0LveqPnMMz7xTHQK3Cm9xwtngicgvQhc2M+3vPD+q8HXnqkxUkLePHXe/1L0yrM8mQiLB+rNQSAqPnN3j02ZihYinmOgcCocArQbkBD84SD25uyrg5EScnpaJeh4uK/G4nx4BFw3S1g+8y6fOHa7D28WoOZeBOQGEu3mToO1B74IKLst0RJyapmWTjxB/qbt2Dkoe5ii8doV86B95p0+8W2IE4/GTEArXAI8Qd4DW6zdNfWmHe65MH5GTsgbP/2N94VgP3MAhUEEdfZLIGifeTefuIYQw34HBxeDJSBjxw3TosR/K4OtuHdtkRNy2UTaX+MhRk+B8dCz94BF4UxQPvOuOHH4xKMw7mhDdwId1vgfmxuH/r77leDfRVLIG2vnvky+8t/ioWfwE6LQGl03i1+Lhrp84lixWeiY4L6ACciwQy5WspkzjYBr7lVdJIVcbjghOF9Ojz57NRgnokPALzHv9IkjOiU6g42WdCcgH3Km041aUn+0+4Vw3kVTyIkFN3evF2lzC1Z6hjMxCq3Va595N584lt0XOgy4L2ACjluF8580XvSN9wOuOmd1kRVymUiLniX8GIm0co5bpE565TOX7pRbX5WLfeATj9QAozHdCWRWch5M2+aK7hfCexdZIZdI0izxiGhLt2BPz/AmSKE1D9bNIkX8llcg4oXyxn3hEeApGXJor3lv6vWBbufWX48jLeS7pl63gxYJrXLA9dcLXIsEgYGKuesThyUeiWFEI/ojIPcWNi2Da+Le/m4L+lqkhVzCoF2ZH2Dtxj5Gy2BxRJ9AsT5z+MSjP6ZoYRcBJ5LOtH7TOHlu6CGHXa2SIdsRP3bWzt4qLHs1TyYj3lI0zyVQqM8cPnGXGH4qQ8Cyha3pS2RkXZTaHHkhl7AsxpfSsn3aRQhWeZQmT39tyXaz5NoDFD7x/ujhWhQJSBevbVnPNpe//3TU2qeEkLdMm/MKxZT/F6zyqE2f/tvjinnP3CzwiffPDVejSYCexwtNiHvYBfW0k1m0DiWEXCKzhLWIrPLDWLYfrQmUrzXZPvO/7NtJ+YVER4ghVmzmY4fr0SHgBFwY5nONQ/auj06rulqilK+iumHJSl6W+rJoD2U3pS5qeFU0gbRtsWOHjGCnDh/H1ux4gyVpKzmlJl/RPcYH4kSAy5WclvFPTbXz10SxX8pY5BKezfj3hGEchK88ilOp/zbJB6DbDh9gj5OIy9cQ8f554Wp0CPBUkgnT2tA0ZUJDdFrVvSVKCfmO2tlbyMeyCr7y7oOoyjudYnClqwUHCChDgCwOipqzuMbvYny6FdV2KyXkEqIl7IW0r+cerPaM6pRCu0AgPgSkNc5Mo6Fx8uxnotwr5YS8Zdq8N4Vt3Y8Ut1GeVmgbCMSAgNyL0zDbbVt8J2px4z3pKifksgNcWIvpgec7yIzYczjxHgRAwCsCjgvXtB/aUTf/j16V6Vc5Sgp5U90N79HGEwvgXvFrWqBcEChxApqTb3y3ZesLVCChpJBLsPoh+yGyyl+Ei0WFaYY2goBaBKSuCCYWtfzTrLdUaLmyQr59+nxaHGTfTpY55dVCMJsKkw1tBAEVCEiXLRmJr7UdaItUhsP+2Ckr5LJTFJz/a0op+V/Ok+X+eolrIAACIFAgAUFP4ei4Y8/0W/YV+JHQb1NayCU9TbPqae+83YxWXuEAARAAgcEQ4GUy3NBc21y797HBlBP0Z5VXv+2T5/+NnFkLEMES9NRBfSAQMwL0gJMZ1j5KU3sr4/W2Sr1TXsglbO2gtYwWCf0BOwmpNPXQVhCIFgH5gNO27bt3TP7Gq9FqWf7WxELI5YNPWiR0I/nL03jwmX/QcQcIgEB3Ak6USnv6z7Y+anH3K2q8i4WQS9TN0+b/N/05vR8PPtWYeGglCESGgIx6s+20ze35LZdceSgy7SqiIbERctln3U59W7S1v4bY8iJmAG4FgRInII0/27Du3zFl3kZVUcRKyLdNu3YPhQ7NF7aN2HJVZyTaDQIBEnBcKm3p15Ls8LcDrNbzqmIl5JJO89Q56+nJ8w/hYvF8rqBAEIgXgYxLxaBEtXO2Tbt1j8qdi52Qy8E4xA7fSdvC/QUuFpWnJtoOAv4SkMYe5Rpf0jR17lP+1uR/6bEU8n3015XcK9+gDSLbEMXi/yRCDSCgGgEZqkwLCTeljTalXSou91gKuexc89S5zwrL/B5cLO5Q4ycIgIBDQGY2tOyDNufXvnfpzQfiQCW2Qi4HJ9UyYgEtFHoGYh6HqYo+gIA3BJykWKZ1547JsyOfZ7zQHsdayN+eMaNNS4hrhGm2MB17RRY6KXAfCMSVAC9LkUvFeKz5UPXSOPUx1kIuB2r7xXPeYKaYSy9t+MvjNHXRFxAojoBjiaeNtzTLmsWmR3cj5eJ6lbk79kIuu9k0bc5PySpfAhfLQKYIPgMCMSAg998UrI3+/7XtdfMbY9Cjbl0oCSGXPeb23m/a7cZGiHm38ccbECgJAjyRkJkN72icMufpOHa4ZIS8qa6+1WLsq8KwGhntAIIDBECgNAjwcvKLtxurm1qfWxTXHpeMkMsB3Fk7eyvFl3+V01csRl+1cIAACMSbgPwGLtrSL1EylevY9J+RLRfPo6SEXA6hXMJvm+ZtHFEs8ZzR6BUIuATod5yeje0yDfHlprob3nNPx/FnyQm5HERaLLTYTqcflF+5cIAACMSQAC36oVQqhjDMmTsvnbs5hj3s1qWSFHIiIMxD++bRV66nZVwpDhAAgRgRIAXntIevSFu3NdfNV2rvzYGOQkk7iiesWVhjJ5NP8YQ2SaTNgTLE50AABCJEwFn0057+QVPtnGsj1Cxfm1KqFrkD1YkntcWXhGntQiSLr/MMhYNAIAScb9jpdAPt43t9IBVGpJKStsjdMahat2QyWeU/Z4INpS2f3NP4CQIgoBABGaFiG8YLQvCpO2pn71Ko6YNuaklb5C49J5LFsq6lkEQTYYkuFfwEAXUIOHsPGObWhOBfKDURl6MEIe+Yqztq5z7ErI6wRLlzCA4QAAElCDg5VGy7xbbSn3+X1ooo0WiPGwkhzwLaNGXu3XbaXOD8dYeWZ5HBSxCIKAEZKy7EXssyv9g87YZNEW2l783CWvUeiA8+/MQzw/73+TH00ORsZsV2IViPXuMtCChIgEIMKYtSK+VQuWLH1HnrFeyBZ02GRd4TJeeiqbJ8nmhPr0KMeU84eA8CESEgF/ww3i4s4+tN0+Y+HpFWhdYMCHku9GfONFIte64R6fb/hJjnAoRzIBAiASniGjcowuya5tp5D4fYkshUDU9wP0MxYfXCCntYgizz5OdoFWg/d+ISCIBAIARIxClEw6QFfNc1T5v3QCB1KlAJLPJ+Bmn79PmHrV3tM+w24+ewzPsBhUsgEAQBKeKcRNywZ0HEuwOHRd6dR8531WuWD2F6249JzC8n33nOe3ASBEDARwLSncJZWpj2dU1T5zzoY01KFg2LvIBha6qb2Zqq2DODpY2fOJY5/vwVQA23gIBHBCg6hUS8nXziX4eI52YKIc/NpdfZty+obxt5YPdX7bSx3NkuDouGejHCCRDwmoBc7EP+lIPCMK6ibdpWeV1+XMpDHHkRI7nrZxutgx84p2Ho+FSFltDPppUI4FcEP9wKAsUQkAvzaLHPbjttf7G5bl5JpKMthk/2vXASZNMo4nVVw+L1WjJ5CSWuL+JTuBUEQKAQAjxFmyVbopF28/oibQTzbCGfKeV7iBaOYgiMb1g6lrb7/B7n/AKIeDHkcC8IFEZAPoei360tmrC+QCIe+919CqPS/12wyPvn0+1q1ZOLz9eEtowlk6cieqUbGrwBAU8IZETceE5Y6Suap97wjieFlkAhEPJCBnlDfaL68Kj5TNe+xTVtCCzxQqDhHhAoggAFDzg73qeNn2kHzGtoDcfuIj5d8rdCyPNMgap19xzD9dQSnkhcSjtyM2aLPJ/AZRAAgaIIZJbc00a61j1V75bftmnmTKOoz+NmWieFo08C1Q1L62g58FKeTE6EK6VPTLgAAgMm4OQSF6yV2eaNlEb6/gEXVOIfhJDnmABOjpWh2p20j+cN9FAzQZZCjrtwCgRAYDAEHFeKZW5jafNrTXXzfz2Yskr9sxDyHjOAolJOpsxqy7Rk4h9Emr7hCbhSeiDCWxAYHAFSHeehZtr4b8No+8quupveGFyB+DQWtGTNgZr1S/6ZHmY+TIt9TnVEPOsaXoIACHhAQCa+oiX3lDPlh+02u/r9afObPSi15IuARU5T4Oi1C0YZ2pDvcl27RqZXw85AJf97AQA+EHC2UBT2XlqpeXPztDlIQesh45IX8sp1i8/R9cQyntTPgCvFw5mFokDAJSBDC8uSjHKIv6RZ1jXbp879g3sJP70hULquFcF49Tn3ztIS2gqyxI+FK8WbCYVSQKAbAZm5kNwptmE/oB8wrtx+2by/dbuON54QKEmL/Jj1i6oMpi/iuv55YdkUG07/cIAACHhKwLHCTWsnt6wbG2vnPuRp4SisG4GSE/LqdYs/yRKJeyl+9UTEhnebC3gDAt4QkAt8KP0sJbxazw0+r6lu1v96UzBK6YtA6Qj56vpU9RGjb+EJ7VZ6oFnurNLsiwrOgwAIDIhAJjbc2s9M6ztN2ysWM6zSHBDHYj9UEtkPx65feEKSkRWeTEx28qRgmX2x8wT3g0D/BKQVnqRHbmlzI4n49U118/7U/wdw1UsCsbfIaYHPZzWdL6KvehNEO1I4eDl5UBYISALSF07ivZeeN93VdHDPUja9HhvbBjw1YivkY371H8OSqfJ/5VybQys1NSyzD3hmobrYE6BgAcYoOT8J+FrKHX779inIHR7WoMdyz86aNd//UKqsYr2WSs2jJfYQ8bBmF+qNJwEZF15Omz8I9jZLWzOafv/+pRDxcIc6dhY5bcH2NbIU/p3+jUZseLiTC7XHj0DHw8w2UvEH9Xa+4N1LZzXFr5fq9Sg2Ql695p4xPJG6m/I4XEUbttIye8SGqzcd0eKoEpDhhIwscXKjPGUK81s7p8x7PqptLcV2xULIq9bRFmwJ/T6WTJzmPNBExsJSnMvosx8E5MrMBO1mb5ivkYgvaD5U9QibPh15nf1gPYgy1Rby1av16qFN8ylvuNyCbSi2YBvETMBHQSCbgAwnpJ3sRbu5QzB7aVmb/YO3L5u3N/sWvI4OAWWFvHrNkqNZUqMt2PRPOREpWGYfnVmFlqhLwN070zD3MsEeFK3t9zV/BpsgR31AlRTymnWLpwldpy3Y9GMRGx71KYb2KUGgU8CNQ9TeR8gfvrDpEiytV2LsqJFKCbncgk0M1e+QW7BRy5OIDVdlmqGdkSVAceC0Jy2lmCUB1/hPNdtailDCyI5Wnw1TRsirn7z3JC7E/SyZvAB5w/scT1wAgcIIdPrAjf1kff+XKaz7dmJBT2HsIniXEkJe9cTiK7im30Ox4ZWIDY/gLEKT1CFAYYRyRSb9Hu2kL+SUWtZ6kHav/6s6HUBLcxGIdNKsiY8tGpmu0L9LFsO/yL84EPFcQ4hzIJCHgPR/yzhwOsgd+To9xFzJ2o3/3/Sp69/N80lcVoRAZC3y8WsWnq2lkssohvVMuFIUmU1oZrQIOBkJKYQwbVpM2L8TCe1H7aZ4fHft7P3RaihaM1gCURRyXt2w+Dqm6XdR7vDhNAkH20d8HgRKh4BrfZOI08Ypuygo4FeMJVY1Tb72d/TNlpY844gjgUgJ+ZhHF1WlhmgLyQq/nFaRYQu2OM449MkfAuT3lu4T+vZqkXjT5sbaI5qR/uX2uvmN/lSIUqNEIDJCXvPEkn+k/ZDvp6Q8H8QWbFGaImhLZAl0uE4oFzg5v+2ttm2v0bhY3Xig5o9YRh/ZUfOlYeELOW3BNv6IUTfrun4bxbGWY5m9L+OMQuNCwMl9Qg8u5S5XltUkBH+KMhE+evhg67N7pt+yLy7dRD+KIxCqkI977HvH6xUV92pJfQq2YCtu4HB3iRCQPm8SbyY3cbDomaVpSVfJRtLxx21T/HbnpXNbSoQEutkPgdCEnDIWfoZ8eovpH7Zg62eAcKkECUiXiRTvzANL+YDydfrfBnpW+QQzzf9pqrvhvRKkgi73QyBwIR+7etkRiWGm3IJtLrZg62dkcKl0CMhl8tLiJvF20k5Y1vsU6/0S5dXfqGniaXOXubnlyhtlDhQcIJCTQKBCXrNm6YdEgi+j9JgfR2x4zvHAybgTIFeJ3Oey0+KWDypp42IKDXxNcPY8F/y3mmlsQrRJ3CeCt/0LTMirG5ZezTR2N4UWYgs2b0wFgX4AAAJSSURBVMcQpUWVgCva5CKR1rZ8QEkGjNy6qomEewtFaW2i/EHPC5He3PzCwXdZfT22tYrqWEa8Xb4LudyCjdEWbGSBYAu2iE8GNG8ABKRYy0Na2VKw6SeJtLSymW2aJrkQm8lF8iade5nu2Cxs4+WkKN+6bdq1ezIfxP9BYPAEfBXy6ieWnEuzWsaGnybSaYp1HXyDUQIIBELA+c2QoixrkyItX9M/+TNzkuYzWdiGSf8Th2hu76JvnI3CZm/RHVuEsF+3uXgjqbPG7ZPn7w6kzaikZAnIWen5Mea5/xiW2lc+n75O1mtDy+Vk97wOFAgCnhBwjAv6n2tkyP1epUBL3zVj7fTvMOl2K12X25y9R78wuwTnjbTYfZuwRaNgVqPQ+btpoe9GDhOJDEcYBHzJfpjcW/Yx2vxhAhf2cvtAK/x+YYws6swQoA0nSZlp2bqM5RPyn03WtUXuaouT84NM7IxQ2yTYTLTSbfvoln1cEwe5nThosfb9bclhe4802tv+VjtbCjsOEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEAABEIgqgf8Dh5Olia/uanYAAAAASUVORK5CYII=',
      extension: 'png',
      options: {stretch: true}
    })

### Sheet.note(col, row, text) 

Add a note to a cell
    
    sheet.note(3, 3, 'Check this out')


Add a note with multiple lines:
  
    sheet.note(3,4, [
      "Twas 'brillig and the slithy tove",
      "Did gyre and gimble in the wabe"
    ]

Add a note with one (or more) lines formatted:

    sheet.note(3,4, [
      "Twas 'brillig and the slithy tove",
      "Did gyre and gimble in the wabe",
      {"text": "All mimsy were the borogroves", { "bold": true, "fontSize": 14, "fontFamily": "Courier"}}
      "And the mome raths outgrabe"
    ]


## Testing

There is a nascent Mocha test suite.
```
> npm test
```

A number of these tests currently writes output files and test for exact matches
against a reference file located in `test/files/`.

It's possible that a future feature extension might
break tests for innocent reasons (e.g. by writing additional XML to the workbook)
in which case, visually inspect the output file and update the reference file.

## Release notes

v1.0.1
Publish to NPM

v0.4.6 
* Async functions return a Promise if no callback is specified, so they can work with async/await syntax
v0.4.4
* Add `workbook.set(data)` to generate an entire workbook as JSON data 

v0.4.3
* Allow notes

v0.4.2
* Extend to allow custom number formats per 

v0.4.1
* Extend Sheet.set() to set multiple rows/columns/cells at once as a dense array of arrays or a nested sparse object.

v0.4.0
* Add images 

v0.3.10
* Add `Sheet.form(ncols, nrows)`

v0.3.9
* Add `Sheet.split(ncols, nrows)`

v0.3.8
* Add `Sheet.sheetViews` and `Sheet.pageSetup(obj)` 

v0.3.3
* Allow more than 26 columns
* Handle Javascript Date objects as cell values, converting to Excel dates
* Added concise form of `.set({})` which now also accepts an object

v0.3.0
*  Port to JSZip 3.1.2 (from 2.5), a breaking change which makes all JSZip methods asynchronous, 

v0.2.0
* Write numbers as numbers
* Write fill and font colors using hex ranges (ported from https://github.com/aloteot/msexcel-builder and applied into coffeescript)
* Apply autofilters
* Add mocha testing


v0.1.0
* Generate JSZip object, dropping need to generate temporary files on disk.
* Removed dependency on `fs-extra` and `exec` and `easy-zip`.
* Added dependency on `js-zip`.
* Removed method `save` and replaced it with `generate(callback)` that returns a JSZip object.
* This now theoretically should be able to run in the browser, though that is not tested.
* Also refactored base Excel files so they are read from code rather than from disk.

v0.0.2:
* Switch compress work to easy-zip to support Heroku deployment.

v0.0.1: Includes

* First release.
* Using 7z.exe to do compress work, so only support windows now.
