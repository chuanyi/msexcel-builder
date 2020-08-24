var fs = require('fs');
var assert = require('assert');
var JSZip = require('jszip');

var excelbuilder = require('..');
var OUTFILE = './test/out/style2.xlsx';
var TESTFILE = './test/files/style.xlsx';

var compareWorkbooks = require('./util/compareworkbooks.js')


describe('It generates a simple workbook with styles applied individuallymo', function () {


  it('generates a ZIP file we can save', function (done) {
    this.timeout(5000)
    var workbook = excelbuilder.createWorkbook()
    var sheet1 = workbook.createSheet('sheet1', 10, 12);

    sheet1.set(1, 1, 'Red bold centered  with border');
    sheet1.set(2, 2, Math.PI);
    sheet1.set(3, 3, '' + Math.PI);
    sheet1.font(1, 1, {
      name: 'Verdana',
      sz: 32,
      color: "FF0022FF",
      bold: true,
      iter: true
    })
    sheet1.align(1, 1, 'center')
    sheet1.fill(1, 1, {
      type: 'solid',
      fgColor: 'FFFF2200'
    })
    sheet1.fill(2, 2, {
      type: 'solid',
      fgColor: 'FF0022FF'
    })
    sheet1.fill(3, 3, {
      type: 'solid',
      fgColor: 'FF22FF00'
    })
    sheet1.numberFormat(2, 2, '0.00%') // 10=>'0.00%'

    // for some reason date formats only work if the fill is set
    sheet1.set(1, 4, new Date('04/01/2009'))
    sheet1.set(1, 5, new Date('04/01/2009'))
    sheet1.fill(1, 5, {
      type: "solid",
      fgColor: "FFAA000"
    })
    sheet1.numberFormat(1, 5, "m/d/yy")


  sheet1.autoFilter(true);
  // Save it
  workbook.generate(function (err, zip) {
    if (err) throw err;
    zip.generateAsync({type: "nodebuffer"}).then(function (buffer) {
      if (err) throw err;
      fs.writeFile(OUTFILE, buffer, function (err) {
        if (err) throw err;
        console.log("open \"" + OUTFILE + "\"")
        compareWorkbooks(TESTFILE, OUTFILE, function (err, result) {
          if (err) throw err;
          done();
        })
      })
    })
  })
})
})
;
