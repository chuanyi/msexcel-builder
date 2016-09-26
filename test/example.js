var fs = require('fs');
var assert = require('assert');
var JSZip = require('jszip');
var compareWorkbooks = require('./util/compareworkbooks.js')

var excelbuilder = require('..');


describe('It generates a simple workbook', function () {

  it('has a vestigial cancel method for backward compatibility', function () {
    var workbook = excelbuilder.createWorkbook()
    workbook.cancel()
  })

  it('generates a ZIP file we can save', function (done) {

    var workbook = excelbuilder.createWorkbook()
    var sheet1 = workbook.createSheet('sheet1', 10, 12);
    sheet1.set(1, 1, 'I am title');
    for (var i = 2; i < 5; i++) {
      sheet1.set(i, 1, 'test' + i);
      sheet1.set(i, 2, i);
    }
    workbook.generate(function (err, zip) {
      if (err) throw err;
        zip.generateAsync({type: "nodebuffer"}).then(function (buffer) {
          var OUTFILE = './test/out/example.xlsx';
          fs.writeFile(OUTFILE, buffer, function (err) {
            console.log('open \"' + OUTFILE + "\"");
            compareWorkbooks('./test/files/example.xlsx', OUTFILE, function (err, result) {
              if (err) throw err;
              assert(result)
              done(err);
            })
          });
        });

    });
  })
  //
  it('Supports the prior constructor syntax', function (done) {
    var PATH = './test/out';
    var FILENAME = 'example2.xlsx';
    var workbook = excelbuilder.createWorkbook(PATH, FILENAME);
    var sheet1 = workbook.createSheet('sheet1', 10, 12);

    sheet1.set(1, 1, 'I am title');
    for (var i = 2; i < 6; i++) {
      sheet1.set(i, 1, 'test' + i);
      sheet1.set(i, 2, i / 2);
    }

    workbook.save(function (err) {
      if (err) throw err;
      else {
        var OUTFILE = PATH + "/" + FILENAME;
        console.log('open \"' + OUTFILE + "\"");        done()
      }
    });
  })

});

