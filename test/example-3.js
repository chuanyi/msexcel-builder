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


  it('takes a file name on save rather than in constructor', function (done) {
    var PATH = './test/out';
    var FILENAME = 'example3.xlsx';
    var workbook = excelbuilder.createWorkbook();
    var sheet1 = workbook.createSheet('sheet1', 10, 12);

    var headerStyle = {

      font: {
        name: 'Arial',
        sz: 14,
        color: "FF000000",
        bold: true,
        iter: false
      },
      align: 'center',
      valign: 'center',
      fill: {
        type: 'solid',
        fgColor: 'FFCCFFCC'
      },
      border: {
        top: "thin",
        bottom: "thin",
        right: "thin",
        left: "thin"
      }
    }


    sheet1.set(1, 1, "WCenter");
    sheet1.set(2, 1, "WRType");
    sheet1.set(3, 1, "Status");
    sheet1.set(4, 1, "Eng");
    sheet1.set(5, 1, "WorkNum");
    sheet1.set(6, 1, "CreationTime");
    sheet1.set(7, 1, "Description");

    sheet1.set(1, 1, headerStyle);
    sheet1.set(2, 1, headerStyle);
    sheet1.set(3, 1, headerStyle);
    sheet1.set(4, 1, headerStyle);
    sheet1.set(5, 1, headerStyle);
    sheet1.set(6, 1, headerStyle);
    sheet1.set(7, 1, headerStyle);

    sheet1.height(1, 40)

    sheet1.width(1, 24)
    sheet1.width(2, 18)
    sheet1.width(3, 18)
    sheet1.width(4, 18)
    sheet1.width(5, 18)
    sheet1.width(6, 18)
    sheet1.width(7, 30)

    sheet1.sheetViews({
      showGridLines: "0",
      zoomScaleNormal: 50,
      zoomScale: 50
    })

    sheet1.pageSetup({
      paperSize: '9',
      orientation: 'landscape',
      horizontalDpi: '200',
      verticalDpi: '200'
    })

    // sheet1.sheetProperties({fitToPage: 1})

    //     "A2": {v: 444483},
    // "B2": {v: "Inventory reconciliation work"},
    // "C2": {v: "Scheduled"},
    // "D2": {v: "V374137"},
    // "E2": {v: null},
    // "F2": {v: new Date("2017-12-17 09:25:39"), t: "d"}

    // sheet1.autoFilter(true);
    // Save it
    var OUTFILE = PATH + "/" + FILENAME;


    workbook.save(OUTFILE, function (err) {
      if (err) throw err;
      else {
        console.log('open \"' + OUTFILE + "\"");
        done()
      }
    });
  })

})
