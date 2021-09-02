var fs = require('fs');
var assert = require('assert');
var JSZip = require('jszip');

function requireUncached(module) {
  delete require.cache[require.resolve(module)];
  return require(module);
}

const excelbuilder = requireUncached('..');
var OUTFILE = './lab/border/borders.xlsx';
var TESTFILE = './test/files/borders.xlsx';

var compareWorkbooks = require('./util/compareworkbooks.js')


describe('It generates a simple workbook with styles applied concisely', function () {


  it('generates a ZIP file we can save', function (done) {
    this.timeout(5000)

    var workbook = excelbuilder.createWorkbook()
    var sheet1 = workbook.createSheet('sheet1', 10, 12);

    sheet1.set(2, 2, {
      set: 'Red borders',
      border: {
          top: {
            style: 'thick',
            color: {
              rgb: 'FFFF0000'
            }
          },
          bottom:{
            style: 'thin',
            color: {
              theme: 5
            }
          },
      },
      align: 'center',
    });


    // Save it
    workbook.generate(function (err, zip) {
      if (err) throw err;
      zip.generateAsync({type: "nodebuffer"}).then(function (buffer) {
        if (err) throw err;
        fs.writeFile(OUTFILE, buffer, function (err) {
          if (err) throw err;
          console.log("open \"" + OUTFILE + "\" ")
          // done(err);

          compareWorkbooks(TESTFILE, OUTFILE, function (err, result) {
            if (!result) return done (new Error("Results don't match"))
            done(err);
          })
        })
      })
    })
  })
});
