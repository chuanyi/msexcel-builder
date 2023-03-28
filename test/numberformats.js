var fs = require('fs');
var assert = require('assert');
var JSZip = require('jszip');
var path = require('path')
var compareWorkbooks = require('./util/compareworkbooks.js')

var excelbuilder = require('../lib/msexcel-builder');


describe('It generates a simple workbook', function () {


  it('takes a file name on save rather than in constructor', function (done) {
    var PATH = './test/out';
    var FILENAME = 'numberformats.xlsx';
    var OUTFILE = PATH + "/" + FILENAME;

    var workbook = excelbuilder.createWorkbook();
    var sheet1 = workbook.createSheet('sheet1', 30, 30);

    var letters = 'abcdefghijklmnopqrstuvwxyz';
    var count=0;

    for (var i=0; i<letters.length; i++) {
      for (var j = 0; j < letters.length; j++) {
        var str = letters[i] + letters[j]
        //var fmt = `"${str}" 0.0%;"${str}" -0.0%;"${str}" –`
        var fmt = `[Blue]"(${count})" 0.0%`

        fmt = (count++ < 206) ? fmt : null

        sheet1.set(1 + i, 1 + j, {
          set: Math.random(),
          numberFormat: fmt
        })
      }
    }


    workbook.save(OUTFILE, function (err) {
      if (err) throw err;
      else {
        console.log('open \"' + OUTFILE + "\"");
        done()
      }
    });
  })

})
