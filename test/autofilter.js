var fs = require('fs');
var assert = require('assert');
var JSZip = require('jszip');

var excelbuilder = require('..');
var OUTFILE = './test/out/autofilter.xlsx';
var TESTFILE = './test/files/autofilter.xlsx';
var compareWorkbooks = require('./util/compareworkbooks.js')


describe('It applies autofilter', function () {


  it('generates a ZIP file we can save', function (done) {

    var workbook = excelbuilder.createWorkbook()

    // Create a new worksheet with 10 columns and 12 rows
    var sheet1 = workbook.createSheet('NEURO RAD', 10, 12);
    var colNames = 'ALPHA,BRAVO,CHARLIE,DELTA,ECHO,FOXTROT,GOLF,HOTEL,INDIA'.split(',');

    for (var c = 0; c < 10; c++) {
      sheet1.set(c + 1, 1, colNames[c]);
    }

    for (var c = 0; c < 10; c++) {
      for (var r = 0; r < 11; r++) {
        sheet1.set(c + 1, r + 2, '' + r * c);
      }
    }

    sheet1.autoFilter(true);

    // Create a new worksheet with 10 columns and 12 rows
    var sheet2 = workbook.createSheet('NEURO ONC', 10, 12);
    var colNames = 'ALPHA,BRAVO,CHARLIE,DELTA,ECHO,FOXTROT,GOLF,HOTEL,INDIA'.split(',');

    for (var c = 0; c < 10; c++) {
      sheet2.set(c + 1, 1, colNames[c]);
    }

    for (var c = 0; c < 10; c++) {
      for (var r = 0; r < 11; r++) {
        sheet2.set(c + 1, r + 2, r * c);
      }
    }

    sheet2.autoFilter('A1:E12');

    workbook.generate(function (err, zip) {
      if (err) throw err;
      else {
        var buffer = zip.generateAsync({type: "nodebuffer"}).then(function (buffer) {
          fs.writeFile(OUTFILE, buffer, function (err) {
            console.log('open \"' + OUTFILE + "\"");
            compareWorkbooks(TESTFILE, OUTFILE, function (err, result) {
              if (!result) return done(new Error("Results don't match"))
              assert(result)
              done();
            })

          });
        })
      }
    });
  })
});
//