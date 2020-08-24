var fs = require('fs');
var assert = require('assert');
var JSZip = require('jszip');
var path = require('path')
var compareWorkbooks = require('./util/compareworkbooks.js')

var excelbuilder = require('..');


describe('It generates a simple workbook', function () {

  it('generates a ZIP file we can save', function (done) {

    var workbook = excelbuilder.createWorkbook();

    var table = [
      [1, 2, "", 4, 5],
      [2, 4, null, 16, 20],
      [1, 4, NaN, 16, 25],
      [4, 8, undefined, 16, 20]
    ]

    var sheet1 = workbook.createSheet('sheet1', table[0].length, table.length);
    table.forEach(function (row, rowIdx) {
      row.forEach(function (val, colIdx) {
        sheet1.set(colIdx + 1, rowIdx + 1, val)
      })
    })

    workbook.generate(function (err, zip) {
      if (err) throw err;
      zip.generateAsync({type: "nodebuffer"}).then(function (buffer) {
        var OUTFILE = './test/out/example.xlsx';
        fs.writeFile(OUTFILE, buffer, function (err) {
          console.log('open \"' + OUTFILE + "\"");
          compareWorkbooks('./test/files/example.xlsx', OUTFILE, function (err, result) {
            if (err) throw err;
            // assert(result)
            done(err);
          });
        });
      });
    });
  });
});

