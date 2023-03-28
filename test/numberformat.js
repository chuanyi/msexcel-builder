var fs = require('fs');
var assert = require('assert');
var JSZip = require('jszip');
var path = require('path')
var compareWorkbooks = require('./util/compareworkbooks.js')

var excelbuilder = require('../lib/msexcel-builder');


describe('It generates a simple workbook', function () {


  it('takes a file name on save rather than in constructor', function (done) {
    var PATH = './test/out';
    var FILENAME = 'numberformat.xlsx';
    var OUTFILE = PATH + "/" + FILENAME;

    //OUTFILE = './lab/format/format.xlsx'

    var workbook = excelbuilder.createWorkbook();
    var sheet1 = workbook.createSheet('sheet1', 10, 12);


    sheet1.set(1, 1, {
      fill: {type: "solid", fgColor: "00FFFFFF"},
      // font : { name: "Calibri", sz: 8 },
      numberFormat: '#,##0',
      set: 1.61803398875
    });
    sheet1.set(1, 2, {
      fill: {type: "solid", fgColor: "00FFFFFF"},
      // font : { name: "Calibri", sz: 8 },
      numberFormat: "0.00",
      set: 2.71828182846
    });
    sheet1.set(1, 3, {
      fill: {type: "solid", fgColor: "00FFFFFF"},
      font: {name: "Calibri", sz: 8},
      numberFormat: "0%",
      set: 3.14159265459
    });

    sheet1.set(1, 4, {
      fill: {type: "solid", fgColor: "00FFFFFF"},
      font: {name: "Calibri", sz: 8},
      numberFormat: "$#,###.00",
      set: 314159.265459
    });
    // sheet1.set(1,3, 3.14159265459)

    // sheet1.numberFormat(1,3, "0%")

    sheet1.set(1, 5, {
      fill: {type: "solid", fgColor: "00FFFFFF"},
      font: {name: "Calibri", sz: 12},
      numberFormat: '" ABCDE "0.0%;" ABCDE "-0.0%;" ABCDE "—;@',
      set: 0.314159
    });

    workbook.save(OUTFILE, function (err) {
      if (err) throw err;
      else {
        console.log('open \"' + OUTFILE + "\"");
        done()
      }
    });
  })

})
