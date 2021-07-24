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

describe('It generates a simple workbook', function () {

  it('has a vestigial cancel method for backward compatibility', function () {
    var workbook = excelbuilder.createWorkbook()
    workbook.cancel()
  })

  it('Simple example #1', function (done) {
    this.timeout(20000)
    var workbook = excelbuilder.createWorkbook()
    var OUTFILE = './test/out/example-1.xlsx';
    var TESTFILE = './test/files/example-1.xlsx'
    var sheet1 = workbook.createSheet('sheet1', 10, 12);
    sheet1.set(1, 1, 'I am title');
    for (var i = 2; i < 5; i++) {
      sheet1.set(i, 1, 'test' + i);
      sheet1.set(i, 2, i);
    }
    workbook.generate(function (err, zip) {
      if (err) throw err;
      else {
        zip.generateAsync({type: "nodebuffer"}).then(function (buffer) {
          fs.writeFile(OUTFILE, buffer, function (err) {
            console.log('open \"' + OUTFILE + "\"");
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
    sheet1.set(2,1, {})

    workbook.save(function (err) {
      if (err) throw err;
      else {
        var OUTFILE = PATH + "/" + FILENAME;
        console.log('open \"' + OUTFILE + "\"");        done()
      }
    });
  })

  it('takes a file name on save rather than in constructor', function (done) {
    var PATH = './test/out';
    var FILENAME = 'example3.xlsx';
    var workbook = excelbuilder.createWorkbook();
    var sheet1 = workbook.createSheet('sheet1', 10, 12);

    sheet1.set(1, 1, 'I am title');
    for (var i = 2; i < 6; i++) {
      sheet1.set(i, 1, 'test' + i);
      sheet1.set(i, 2, i / 2);
    }
    var OUTFILE = PATH + "/" + FILENAME;


    workbook.save(OUTFILE, function (err) {
      if (err) throw err;
      else {
        console.log('open \"' + OUTFILE + "\"");
        done()
      }
    });
  })

});

