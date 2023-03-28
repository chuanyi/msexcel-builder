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
    var sheet = workbook.createSheet('sheet1', 53, 27);

    // sheet.set(rows)

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
          if (cell.set &&! Number.isNaN(cell.set)) {
            sheet.set(colIdx + 1, rowIdx + 1, cell.set)
          }
        }
        else if (cell && cell !== undefined && cell !== '' && !Number.isNaN(cell)) {
          sheet.set(colIdx + 1, rowIdx + 1, ["B",colIdx+1,rowIdx+1].join(":"))
        }
      })
    })


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



var rows = [
  [
    null,
    null,
    {},
    {
      "set": "OVERALL",
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      },
      "align": "center"
    },
    null,
    {},
    {
      "set": "A5InteriorFruits_1",
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      },
      "align": "center"
    },
    null,
    null,
    null,
    null,
    {},
    {
      "set": "A5InteriorFruits_2",
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      },
      "align": "center"
    },
    null,
    null,
    null,
    null,
    {},
    {
      "set": "A5InteriorFruits_3",
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      },
      "align": "center"
    },
    null,
    null,
    null,
    null,
    {},
    {
      "set": "A5InteriorFruits_4",
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      },
      "align": "center"
    },
    null,
    null,
    null,
    null,
    {},
    {
      "set": "A5InteriorFruits_5",
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      },
      "align": "center"
    },
    null,
    null,
    null,
    null,
    {},
    {
      "set": "A5InteriorFruits_6",
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      },
      "align": "center"
    },
    null,
    null,
    null,
    null,
    {},
    {
      "set": "A5InteriorFruits_7",
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      },
      "align": "center"
    },
    null,
    null,
    null,
    null,
    {},
    {
      "set": "A5InteriorFruits_8",
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      },
      "align": "center"
    },
    null,
    null,
    null,
    null,
    {}
  ],
  [
    null,
    null,
    null,
    {
      "set": "Overall",
      "font": {
        "bold": true,
        "iter": "-",
        "sz": "11",
        "color": "-",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF",
        "bgColor": "-"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "Total",
      "font": {
        "bold": true,
        "iter": "-",
        "sz": "11",
        "color": "-",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF",
        "bgColor": "-"
      },
      "align": "center",
      "wrap": true
    },
    null,
    {
      "set": "Overall",
      "font": {
        "bold": true,
        "iter": "-",
        "sz": "11",
        "color": "-",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF",
        "bgColor": "-"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "Not an area of strength",
      "font": {
        "bold": true,
        "iter": "-",
        "sz": "11",
        "color": "-",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF",
        "bgColor": "-"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "I am doing okay in this area",
      "font": {
        "bold": true,
        "iter": "-",
        "sz": "11",
        "color": "-",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF",
        "bgColor": "-"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "This is an area of strength",
      "font": {
        "bold": true,
        "iter": "-",
        "sz": "11",
        "color": "-",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF",
        "bgColor": "-"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "This is an area of great strength",
      "font": {
        "bold": true,
        "iter": "-",
        "sz": "11",
        "color": "-",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF",
        "bgColor": "-"
      },
      "align": "center",
      "wrap": true
    },
    null,
    {
      "set": "Overall",
      "font": {
        "bold": true,
        "iter": "-",
        "sz": "11",
        "color": "-",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF",
        "bgColor": "-"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "Not an area of strength",
      "font": {
        "bold": true,
        "iter": "-",
        "sz": "11",
        "color": "-",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF",
        "bgColor": "-"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "I am doing okay in this area",
      "font": {
        "bold": true,
        "iter": "-",
        "sz": "11",
        "color": "-",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF",
        "bgColor": "-"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "This is an area of strength",
      "font": {
        "bold": true,
        "iter": "-",
        "sz": "11",
        "color": "-",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF",
        "bgColor": "-"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "This is an area of great strength",
      "font": {
        "bold": true,
        "iter": "-",
        "sz": "11",
        "color": "-",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF",
        "bgColor": "-"
      },
      "align": "center",
      "wrap": true
    },
    null,
    {
      "set": "Overall",
      "font": {
        "bold": true,
        "iter": "-",
        "sz": "11",
        "color": "-",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF",
        "bgColor": "-"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "Not an area of strength",
      "font": {
        "bold": true,
        "iter": "-",
        "sz": "11",
        "color": "-",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF",
        "bgColor": "-"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "I am doing okay in this area",
      "font": {
        "bold": true,
        "iter": "-",
        "sz": "11",
        "color": "-",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF",
        "bgColor": "-"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "This is an area of strength",
      "font": {
        "bold": true,
        "iter": "-",
        "sz": "11",
        "color": "-",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF",
        "bgColor": "-"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "This is an area of great strength",
      "font": {
        "bold": true,
        "iter": "-",
        "sz": "11",
        "color": "-",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF",
        "bgColor": "-"
      },
      "align": "center",
      "wrap": true
    },
    null,
    {
      "set": "Overall",
      "font": {
        "bold": true,
        "iter": "-",
        "sz": "11",
        "color": "-",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF",
        "bgColor": "-"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "Not an area of strength",
      "font": {
        "bold": true,
        "iter": "-",
        "sz": "11",
        "color": "-",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF",
        "bgColor": "-"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "I am doing okay in this area",
      "font": {
        "bold": true,
        "iter": "-",
        "sz": "11",
        "color": "-",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF",
        "bgColor": "-"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "This is an area of strength",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "This is an area of great strength",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    },
    null,
    {
      "set": "Overall",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "Not an area of strength",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "I am doing okay in this area",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "This is an area of strength",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "This is an area of great strength",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    },
    null,
    {
      "set": "Overall",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "Not an area of strength",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "I am doing okay in this area",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "This is an area of strength",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "This is an area of great strength",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    },
    null,
    {
      "set": "Overall",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "Not an area of strength",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "I am doing okay in this area",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "This is an area of strength",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "This is an area of great strength",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    },
    null,
    {
      "set": "Overall",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "Not an area of strength",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "I am doing okay in this area",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "This is an area of strength",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    },
    {
      "set": "This is an area of great strength",
      "font": {
        "bold": true
      },
      "fill": {
        "type": "solid",
        "fgColor": "FFBFBFBF"
      },
      "align": "center",
      "wrap": true
    }
  ],
  [],
  [
    {
      "set": "OVERALL",
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    }
  ],
  [
    {
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-",
        "right": "-"
      }
    },
    {
      "set": "n-size",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-"
      }
    },
    {
      "set": " "
    },
    {
      "set": 3137,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 3137,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 3137,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 333,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 1061,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 1174,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 569,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 3137,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 289,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 1392,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 1122,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 334,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 3137,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 276,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 1296,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 1175,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 390,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 3137,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 103,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 650,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 1377,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 1007,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 3137,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 154,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 974,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 1349,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 660,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 3137,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 104,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 1023,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 1424,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 586,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 3137,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 359,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 1606,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 972,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 200,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 3137,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 512,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 1728,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 744,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 153,
      "align": "center",
      "numberFormat": "0",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null
  ],
  [],
  [
    {
      "set": "A2. Which of the following have had the most meaningful impact on your faith life? Select all that apply. - Selected Choice ",
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF",
        "iter": "-",
        "sz": "11",
        "name": "Calibri",
        "scheme": "minor",
        "family": "2",
        "underline": "-",
        "strike": "-",
        "outline": "-",
        "shadow": "-"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE",
        "bgColor": "-"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    },
    {
      "font": {
        "bold": true,
        "color": "FFFFFFFF"
      },
      "fill": {
        "type": "solid",
        "fgColor": "FF009CDE"
      }
    }
  ],
  [
    {
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-",
        "right": "-"
      }
    },
    {
      "set": "A2_1: Bible study",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-"
      }
    },
    {
      "set": " "
    },
    {
      "set": 0.5259802358941664,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5259802358941664,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.5259802358941664,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5135135135135135,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5146088595664468,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5357751277683135,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5342706502636204,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.5259802358941664,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4809688581314879,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5258620689655172,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5383244206773619,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5239520958083832,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.5259802358941664,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4927536231884058,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.529320987654321,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5327659574468085,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.517948717948718,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.5259802358941664,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5048543689320388,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5030769230769231,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5453885257806826,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5163853028798411,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.5259802358941664,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.487012987012987,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5143737166324436,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5174203113417346,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5696969696969697,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.5259802358941664,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4423076923076923,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5083088954056696,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5428370786516854,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5307167235494881,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.5259802358941664,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.46518105849582175,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5236612702366127,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5473251028806584,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.55,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.5259802358941664,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.529296875,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5353009259259259,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5174731182795699,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.45098039215686275,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "medium",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null
  ],
  [
    {
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-",
        "right": "-"
      }
    },
    {
      "set": "A2_2: Small group / ministry / community",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-"
      }
    },
    {
      "set": " "
    },
    {
      "set": 0.503984698756774,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.503984698756774,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.503984698756774,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.34534534534534533,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.470311027332705,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5596252129471891,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5448154657293497,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.503984698756774,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4117647058823529,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4885057471264368,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5445632798573975,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5119760479041916,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.503984698756774,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.40942028985507245,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5169753086419753,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5072340425531915,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.517948717948718,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.503984698756774,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4077669902912621,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5169230769230769,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.49963689179375453,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5114200595829196,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.503984698756774,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.44155844155844154,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.48254620123203285,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5233506300963677,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5106060606060606,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.503984698756774,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.47115384615384615,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4965786901270772,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5168539325842697,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.49146757679180886,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.503984698756774,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.467966573816156,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5049813200498132,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5246913580246914,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.46,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.503984698756774,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4609375,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5266203703703703,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4959677419354839,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.43137254901960786,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null
  ],
  [
    {
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-",
        "right": "-"
      }
    },
    {
      "set": "A2_3: Parent",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-"
      }
    },
    {
      "set": " "
    },
    {
      "set": 0.448836467963022,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.448836467963022,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.448836467963022,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.44144144144144143,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.44015080113100846,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.46337308347529815,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.43936731107205623,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.448836467963022,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4083044982698962,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4454022988505747,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.46345811051693403,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4491017964071856,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.448836467963022,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.42391304347826086,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4382716049382716,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.465531914893617,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4512820512820513,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.448836467963022,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3592233009708738,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.43846153846153846,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4466230936819172,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.46772591857000995,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.448836467963022,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.461038961038961,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.41786447638603696,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.469236471460341,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.45,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.448836467963022,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4423076923076923,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4095796676441838,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.46769662921348315,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4726962457337884,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.448836467963022,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.45125348189415043,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4302615193026152,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.47016460905349794,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.49,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.448836467963022,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.42578125,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4519675925925926,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.45564516129032256,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.45751633986928103,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null
  ],
  [
    {
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-",
        "right": "-"
      }
    },
    {
      "set": "A2_4: Non-parent family member",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-"
      }
    },
    {
      "set": " "
    },
    {
      "set": 0.1858463500159388,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1858463500159388,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.1858463500159388,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.17117117117117117,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1828463713477851,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.19420783645655879,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1827768014059754,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.1858463500159388,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1695501730103806,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1853448275862069,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.18983957219251338,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.18862275449101795,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.1858463500159388,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.18478260869565216,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1882716049382716,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.18553191489361703,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1794871794871795,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.1858463500159388,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.11650485436893204,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.20923076923076922,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1924473493100944,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.16881827209533268,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.1858463500159388,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.21428571428571427,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.19815195071868583,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.16604892512972572,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.2015151515151515,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.1858463500159388,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.17307692307692307,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.18377321603128055,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.19803370786516855,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1621160409556314,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.1858463500159388,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1977715877437326,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.17496886674968867,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.20164609053497942,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.175,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.1858463500159388,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1796875,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1892361111111111,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1935483870967742,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.13071895424836602,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null
  ],
  [
    {
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-",
        "right": "-"
      }
    },
    {
      "set": "A2_5: Friend",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-"
      }
    },
    {
      "set": " "
    },
    {
      "set": 0.3736053554351291,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3736053554351291,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.3736053554351291,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.2702702702702703,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.353440150801131,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.41567291311754684,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.38488576449912126,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.3736053554351291,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.328719723183391,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3706896551724138,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3948306595365419,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3532934131736527,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.3736053554351291,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.30434782608695654,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.38503086419753085,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.38468085106382977,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.35128205128205126,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.3736053554351291,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3106796116504854,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.37538461538461537,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3769063180827887,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3743793445878848,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.3736053554351291,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.33766233766233766,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.39117043121149897,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.36619718309859156,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3712121212121212,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.3736053554351291,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3557692307692308,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3782991202346041,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3714887640449438,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.37372013651877134,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.3736053554351291,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3342618384401114,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3823163138231631,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.39094650205761317,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.29,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.3736053554351291,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.369140625,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.38425925925925924,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.35618279569892475,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.35294117647058826,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null
  ],
  [
    {
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-",
        "right": "-"
      }
    },
    {
      "set": "A2_6: Teacher",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-"
      }
    },
    {
      "set": " "
    },
    {
      "set": 0.15301243226012112,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.15301243226012112,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.15301243226012112,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.0990990990990991,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.12252591894439209,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1643952299829642,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.2179261862917399,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.15301243226012112,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.12802768166089964,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.14080459770114942,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.16131907308377896,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.19760479041916168,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.15301243226012112,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.13768115942028986,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1427469135802469,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1523404255319149,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.2,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.15301243226012112,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.0970873786407767,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.15384615384615385,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.14960058097313,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1628599801390268,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.15301243226012112,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.11038961038961038,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.14784394250513347,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.15048183839881393,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.17575757575757575,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.15301243226012112,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.15384615384615385,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.14271749755620725,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.15098314606741572,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1757679180887372,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.15301243226012112,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.12813370473537605,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.14881693648816938,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.16049382716049382,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.195,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.15301243226012112,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.138671875,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.15046296296296297,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.16129032258064516,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1895424836601307,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null
  ],
  [
    {
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-",
        "right": "-"
      }
    },
    {
      "set": "A2_7: Print and/or digital media (e.g., book, podcast)",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-"
      }
    },
    {
      "set": " "
    },
    {
      "set": 0.5890978642014664,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5890978642014664,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.5890978642014664,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6156156156156156,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6079170593779454,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5894378194207837,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5377855887521968,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.5890978642014664,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5847750865051903,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6012931034482759,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5730837789661319,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5958083832335329,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.5890978642014664,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5398550724637681,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6095679012345679,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5897872340425532,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5538461538461539,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.5890978642014664,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5145631067961165,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6030769230769231,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5882352941176471,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5888778550148958,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.5890978642014664,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6233766233766234,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6149897330595483,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5774647887323944,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5666666666666667,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.5890978642014664,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5673076923076923,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5806451612903226,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5990168539325843,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5836177474402731,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.5890978642014664,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6239554317548747,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5815691158156912,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5997942386831275,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.535,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.5890978642014664,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.62890625,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5983796296296297,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5524193548387096,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5294117647058824,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null
  ],
  [
    {
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-",
        "right": "-"
      }
    },
    {
      "set": "A2_8: Prayer",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-"
      }
    },
    {
      "set": " "
    },
    {
      "set": 0.8138348740835193,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.8138348740835193,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.8138348740835193,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7567567567567568,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.8294062205466541,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.823679727427598,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7978910369068541,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.8138348740835193,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7577854671280276,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.8182471264367817,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.8163992869875223,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.8353293413173652,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.8138348740835193,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7282608695652174,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.8047839506172839,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.843404255319149,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.8153846153846154,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.8138348740835193,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.8446601941747572,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.8215384615384616,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.8126361655773421,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.8073485600794439,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.8138348740835193,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7857142857142857,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.797741273100616,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.816160118606375,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.8393939393939394,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.8138348740835193,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7692307692307693,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.8103616813294232,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.8230337078651685,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.8054607508532423,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.8138348740835193,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7910863509749304,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.811332503113325,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.8261316872427984,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.815,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.8138348740835193,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.826171875,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.8125,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.8104838709677419,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.803921568627451,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null
  ],
  [
    {
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-",
        "right": "-"
      }
    },
    {
      "set": "A2_9: Clergy / Religious",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-"
      }
    },
    {
      "set": " "
    },
    {
      "set": 0.6362766974816704,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6362766974816704,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.6362766974816704,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5405405405405406,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6154571159283695,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6780238500851788,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6449912126537786,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.6362766974816704,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.615916955017301,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6271551724137931,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6515151515151515,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6407185628742516,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.6362766974816704,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5652173913043478,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6388888888888888,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6502127659574468,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6358974358974359,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.6362766974816704,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5728155339805825,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6507692307692308,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6289034132171387,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6434955312810328,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.6362766974816704,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5584415584415584,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6303901437371663,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6493699036323203,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6363636363636364,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.6362766974816704,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5865384615384616,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6011730205278593,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6601123595505618,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6484641638225256,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.6362766974816704,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6267409470752089,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.638854296388543,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6481481481481481,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.575,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.6362766974816704,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.634765625,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6412037037037037,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6344086021505376,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5947712418300654,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null
  ],
  [
    {
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-",
        "right": "-"
      }
    },
    {
      "set": "A2_10: Event / Encounter",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-"
      }
    },
    {
      "set": " "
    },
    {
      "set": 0.3704175964297099,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3704175964297099,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.3704175964297099,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.2972972972972973,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.352497643732328,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3628620102214651,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.46221441124780316,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.3704175964297099,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3494809688581315,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.33620689655172414,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.39572192513368987,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.44610778443113774,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.3704175964297099,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3442028985507246,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3464506172839506,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3931914893617021,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.3704175964297099,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3786407766990291,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.33692307692307694,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.37254901960784315,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.38828202581926513,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.3704175964297099,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4090909090909091,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3490759753593429,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.374351371386212,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.38484848484848483,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.3704175964297099,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3942307692307692,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.36950146627565983,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.36235955056179775,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3873720136518771,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.3704175964297099,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3788300835654596,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.35678704856787047,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.38580246913580246,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.39,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.3704175964297099,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.373046875,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3680555555555556,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3844086021505376,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3202614379084967,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null
  ],
  [
    {
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-",
        "right": "-"
      }
    },
    {
      "set": "A2_11: Class / Talk",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-"
      }
    },
    {
      "set": " "
    },
    {
      "set": 0.2843481032833918,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.2843481032833918,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.2843481032833918,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.24324324324324326,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.25164938737040526,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.2887563884156729,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3602811950790861,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.2843481032833918,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.2629757785467128,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.27729885057471265,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.2798573975044563,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.3473053892215569,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.2843481032833918,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.2427536231884058,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.2847222222222222,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.2919148936170213,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.28974358974358977,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.2843481032833918,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.30097087378640774,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.26615384615384613,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.29121278140885987,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.28500496524329694,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.2843481032833918,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.2987012987012987,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.2669404517453799,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.29132690882134915,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.2924242424242424,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.2843481032833918,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.23076923076923078,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.27663734115347016,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.28441011235955055,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.30716723549488056,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.2843481032833918,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.2785515320334262,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.28393524283935245,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.28703703703703703,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.285,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.2843481032833918,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.263671875,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.30324074074074076,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.25806451612903225,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.2679738562091503,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null
  ],
  [
    {
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-",
        "right": "-"
      }
    },
    {
      "set": "A2_12: Sacrament",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-"
      }
    },
    {
      "set": " "
    },
    {
      "set": 0.7593241950908511,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7593241950908511,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.7593241950908511,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6666666666666666,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7511781338360037,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7913117546848382,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7627416520210897,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.7593241950908511,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6955017301038062,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7341954022988506,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7941176470588235,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.8023952095808383,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.7593241950908511,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.6956521739130435,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7337962962962963,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.796595744680851,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7769230769230769,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.7593241950908511,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7766990291262136,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7876923076923077,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7458242556281772,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7576961271102284,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.7593241950908511,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7337662337662337,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7351129363449692,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7679762787249814,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7833333333333333,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.7593241950908511,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7307692307692307,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.729227761485826,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7710674157303371,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.78839590443686,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.7593241950908511,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7325905292479109,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7615193026151931,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7664609053497943,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.755,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.7593241950908511,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.763671875,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7604166666666666,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7486559139784946,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.7843137254901961,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null
  ],
  [
    {
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-",
        "right": "-"
      }
    },
    {
      "set": "A2_13: Lives of the saints",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-"
      }
    },
    {
      "set": " "
    },
    {
      "set": 0.4430985017532674,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4430985017532674,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.4430985017532674,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4084084084084084,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.41470311027332707,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.44718909710391824,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.507908611599297,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.4430985017532674,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.41522491349480967,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.41020114942528735,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4696969696969697,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5149700598802395,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.4430985017532674,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.40217391304347827,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4158950617283951,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.46297872340425533,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.5025641025641026,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.4430985017532674,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.49514563106796117,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4323076923076923,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.439360929557008,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.44985104270109233,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.4430985017532674,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.37012987012987014,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4271047227926078,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4425500370644922,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.48484848484848486,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.4430985017532674,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.36538461538461536,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.41055718475073316,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.45997191011235955,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4726962457337884,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.4430985017532674,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.42896935933147634,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.43711083437110837,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.44753086419753085,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.495,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.4430985017532674,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.453125,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.44560185185185186,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.4260752688172043,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.46405228758169936,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null
  ],
  [
    {
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-",
        "right": "-"
      }
    },
    {
      "set": "A2_14: Other (specify)",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-"
      }
    },
    {
      "set": " "
    },
    {
      "set": 0.14026139623844439,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.14026139623844439,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.14026139623844439,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.12312312312312312,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.11969839773798303,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.14821124361158433,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.17223198594024605,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.14026139623844439,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.11072664359861592,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.15373563218390804,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.13101604278074866,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1407185628742515,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.14026139623844439,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1956521739130435,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.12654320987654322,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.13617021276595745,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.15897435897435896,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.14026139623844439,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1650485436893204,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.12,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.13870733478576616,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1529294935451837,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.14026139623844439,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.12337662337662338,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.12833675564681724,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.13936249073387694,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.16363636363636364,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.14026139623844439,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.15384615384615385,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.13978494623655913,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1306179775280899,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1621160409556314,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.14026139623844439,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.13649025069637882,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.13138231631382316,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.15020576131687244,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.17,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.14026139623844439,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1640625,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.13541666666666666,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.1303763440860215,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.16339869281045752,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null
  ],
  [
    {
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": "-",
        "right": "-"
      }
    },
    {
      "set": "A2_15: Nothing has had a meaningful impact on my faith life",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "left": "-"
      }
    },
    {
      "set": " "
    },
    {
      "set": 0.000318775900541919,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.000318775900541919,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.000318775900541919,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.0008517887563884157,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.000318775900541919,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.00089126559714795,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.000318775900541919,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.000851063829787234,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.000318775900541919,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.0015384615384615385,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.000318775900541919,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.001026694045174538,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.000318775900541919,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.0007022471910112359,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.000318775900541919,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.0006226650062266501,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null,
    {
      "set": 0.000318775900541919,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0.0013440860215053765,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    {
      "set": 0,
      "align": "center",
      "numberFormat": "0%;0%;–;@",
      "border": {
        "top": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "bottom": {
          "style": "thin",
          "color": {
            "rgb": "FFA6A6A6"
          }
        },
        "left": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        },
        "right": {
          "style": "thin",
          "color": {
            "rgb": "FFD9D9D9"
          }
        }
      },
      "note": null
    },
    null
  ],
  [
    {
      "set": "Compact to Yes"
    }
  ]
]
