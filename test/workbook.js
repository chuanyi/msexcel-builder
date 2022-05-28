var fs = require('fs');
var assert = require('assert');
var JSZip = require('jszip');
var path = require('path')
var util = require('util')
var compareWorkbooks = require('./util/compareworkbooks.js')

function requireUncached(module) {
  delete require.cache[require.resolve(module)];
  return require(module);
}

const excelbuilder = requireUncached('..');

describe('It sets cells ', function () {


  it('Generates workbook from JSON', async function () {

    var OUTFILE = './test/out/workbook.xlsx';
    var TESTFILE = './test/files/workbook.xlsx'

    var three = {
      set: 3,
      font: {
        name: 'Verdana',
        sz: 32,
        color: "FF0022FF",
        bold: true,
        iter: true,
        underline: true
      },
      align: 'center',
      fill: {
        type: 'solid',
        fgColor: 'FFFF2200'
      },
      width: 100,
      row:100
    }


    var pojo = {
      "worksheets": [
        {
          "name": "sheet1",
          "cells": [
            ["A", "B", "C", "D", "E"],
              [],
            [1, 2, three, 4, 5],
            [6, 7, 8, 9, 10],
            [11, 12, 13, 14, 15],

          ],
          options: {
            "sheetViews": {
              "showGridLines": "0"
            }
          }
        },
        {
          "name": "sheet2",
          "cells": [
            ["A", "B", "C", "D", "E"],
            [1, 2, three, 4, 5],
            [6, 7, 8, 9, 10],
            [11, 12, 13, 14, 15],
          ]
        }
      ]
    }


    var workbook = excelbuilder.createWorkbook()
    workbook.set(pojo)
    var zip = await workbook.generate()
    var buffer = await zip.generateAsync({type: "nodebuffer"})
    fs.writeFileSync(OUTFILE, buffer)
    console.log("open \"" + OUTFILE + "\"")
    var result = await compareWorkbooks(TESTFILE, OUTFILE)
    if (!result) throw new Error(["Results don't match #1", TESTFILE, OUTFILE].join(":"))
    else return true
  })
})
;

