var fs = require('fs');
var assert = require('assert');
var JSZip = require('jszip');
var path = require('path')
var _ = require('lodash')
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
    var OUTFILE = './test/out/from_data.xlsx';
    var TESTFILE = './test/files/from_data.xlsx'
    var sheet = workbook.createSheet('sheet1', 53, 53);

    _.each(rows, function (row, rowIdx) {
      _.each(row, function (cell, colIdx) {
        var val = (cell && typeof cell == 'object' && 'set' in cell) ? cell.set : cell
        if (cell && typeof cell == 'object') {
          if (cell.colspan) {
            sheet.merge(
                {col: colIdx + 1, row: rowIdx + 1},
                {col: colIdx + cell.colspan, row: rowIdx + 1}
            )
            delete cell.colspan
          }
          if (cell.width) {
            sheet.width(colIdx + 1, cell.width)
          }
          delete cell.width
        }

        if (cell && typeof cell == 'object') {
          if (cell.set === '' || Number.isNaN(cell.set)) delete cell.set
          sheet.set(colIdx + 1, rowIdx + 1, cell.set)
        }
        else if (cell !== undefined && cell !== '' && !Number.isNaN(cell)) {
          sheet.set(colIdx + 1, rowIdx + 1, cell)
        }
      })
    })


    workbook.generate(function (err, zip) {
      if (err) throw err;
      else {
        zip.generateAsync({type: "nodebuffer"}).then(function (buffer) {
          fs.writeFile(OUTFILE, buffer, function (err) {
            console.log('open \"' + OUTFILE + "\"");
            return done()
            compareWorkbooks(TESTFILE, OUTFILE, function (err, result) {
              if (!result) {
                return done(new Error(["Results don't match #1", TESTFILE, OUTFILE].join(":")))
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


var rows = [

  _.range(0, 51).map(v => 'COL' + v),
  _.range(0, 51).map(Math.random),
  _.range(0, 51).map(Math.random),
  _.range(0, 51).map(Math.random),
  _.range(0, 51).map(Math.random)
]
