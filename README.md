# msexcel-builder

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
npm install msexcel-builder-colorfix
```

```javascript
var excelbuilder = require('msexcel-builder-colorfix');
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
      workbook.cancel();
    else
      console.log('congratulations, your workbook created');
  });
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

Notes: The sheet name must be unique within a same workbook.

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

### Sheet.set(col, row, str)

Set the cell data.

* `col` - (Number) Cell column index(start with 1).
* `row` - (Number) Cell row index(start with 1).
* `str` - (String) Cell data.

No returns.

Example:

```javascript
sheet1.set(1,1,'Hello ');
sheet1.set(2,1,'world!');
```

## Sheet.width(col, width)
## Sheet.height(row, height)

Set the column width or row height

Example:

```javascript
sheet1.width(1, 30);
sheet1.height(1, 20);
```

## Sheet.align(col, row, align)
## Sheet.valign(col, row, valign)
## Sheet.wrap(col, row, wrap)
## Sheet.rotate(col, row, angle)

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

## Sheet.font(col, row, font_style)
## Sheet.fill(col, row, fill_style)
## Sheet.border(col, row, border_style)

Set cell font style, fill style or border style

* `font_style` - (Object) font style options 
The options may contain:

  * `name` - (String) font name
  * `sz` - (String) font size
  * `family` - (String) font family
  * `scheme` - (String) font scheme
  * `bold` - (String) if bold: 'true'/'false'
  * `iter` - (String) if italic: 'true'/'false'

* `fill_style` - (Object) fill style options
The options may contain:

  * `type` - (String) fill type: such as 'solid'
  * `fgColor` - (String) front color
  * `bgColor` - (String) background color

* `border_style` - (Object) border style options
The options may contain:

  * `left` - (String) style: 'thin'/'medium'/'thick'/'double'
  * `top` - (String) style: 'thin'/'medium'/'thick'/'double'
  * `right` - (String) style: 'thin'/'medium'/'thick'/'double'
  * `bottom` - (String) style: 'thin'/'medium'/'thick'/'double'

Example:

```javascript
sheet1.font(2, 1, {name:'黑体',sz:'24',family:'3',scheme:'-',bold:'true',iter:'true'});
sheet1.fill(3, 3, {type:'solid',fgColor:'FFFF0000',bgColor:'64'});
sheet1.border(1, 1, {left:'medium',top:'medium',right:'thin',bottom:'medium'});
```

## Sheet.merge(from_cell, to_cell)

Merge some cell ranges

* `from_cell` / `to_cell` - (Object) cell position
The cell object contains:

  * `col` - (Number) cell column index(start with 1)
  * `row` - (Number) cell row index(start with 1) 

Example: Merge the first row as title from (1,1) to (5,1)

```javascript
sheet1.merge({col:1,row:1},{col:5,row:1});
```

## Testing

In node.js

```
> cd test
> node test.js
```

## Release notes

v0.0.2:
* Switch compress work to easy-zip to support Heroku deployment.

v0.0.1: Includes

* First release
* Using 7z.exe to do compress work, so only support windows now.
