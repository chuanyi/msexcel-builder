var fs = require('fs');
var assert = require('assert');
var JSZip = require('jszip');

var excelbuilder = require('..');
var OUTFILE = './test/out/style2.xlsx';
var TESTFILE = './test/files/style.xlsx';

var compareWorkbooks = require('./util/compareworkbooks.js')


describe('It generates a simple workbook', function () {


  it('generates a ZIP file we can save', function (done) {

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
    sheet1.autoFilter(true);
    // Save it
    workbook.generate(function (err, zip) {
      if (err) throw err;
      zip.generateAsync({type: "nodebuffer"}).then(function (buffer) {
        if (err) throw err;
        console.log("Done...")
        fs.writeFile(OUTFILE, buffer, function (err) {
          if (err) throw err;
          compareWorkbooks(TESTFILE, OUTFILE, function (err, result) {
            assert(result)
            done(err);
          })
        })
      })
    })
  })
});
