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

describe('Notes', function () {

  it('adds notes', function (done) {
    var PATH = './test/out';
    var FILENAME = 'notes.xlsx';

    // for local testing
    var PATH = './lab/generated'
    var FILENAME = 'generated.xlsx';

    var workbook = excelbuilder.createWorkbook();
    var sheet1 = workbook.createSheet('sheet1', 5, 2)



    sheet1.set(1, 1, 'I am title');
    for (var i = 2; i < 6; i++) {
      sheet1.set(i, 1, 'test' + i);
      sheet1.set(i, 2, i / 2);

    }
    sheet1.note(2,2, `Apple`)
    sheet1.note(3,2, `Berry`)
    sheet1.note(4,2, `Cherry`)
    sheet1.note(5,2, `Date`)
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

