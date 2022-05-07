var fs = require('fs');
var assert = require('assert');
var JSZip = require('jszip');
var path = require('path')
var compareWorkbooks = require('./util/compareworkbooks.js')

function requireUncached(module) {
  delete require.cache[require.resolve(module)];
  return require(module);
}

const excelbuilder = requireUncached('..');

describe('It sets cells ', function () {


  it('It sets cells', function (done) {
    this.timeout(20000)
    var workbook = excelbuilder.createWorkbook()
    var OUTFILE = './test/out/set.xlsx';
    var TESTFILE = './test/files/set.xlsx'
    var sheet1 = workbook.createSheet('sheet1', 10, 12);


    sheet1.set(1, 1, 'Red');
    sheet1.set(1,2, { set: "Red", font: { bold: true, size: 14, color: "FF2244"}})
    sheet1.set({
      "2": {
        "4":{ set: "B4", font: { bold: true, size: 14, color: "44FF22"}}
      }
    })

    // note that arrays are zero-based indexes, and Worksheet is 1-based
    sheet1.set([null,null,null, [null, "C1"]])

    workbook.generate(function (err, zip) {
      if (err) throw err;
      else {
        zip.generateAsync({type: "nodebuffer"}).then(function (buffer) {
          fs.writeFile(OUTFILE, buffer, function (err) {
            console.log('open \"' + OUTFILE + "\"");
            //return done()
            compareWorkbooks(TESTFILE, OUTFILE, function (err, result) {
              if (!result) {
                return done(new Error(["Results don't match #1",TESTFILE,OUTFILE].join(":")))
              }
              else {
                return done();
              }
            })
          })
        })
      }
    })
  })

});

