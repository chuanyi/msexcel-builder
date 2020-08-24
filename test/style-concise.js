var fs = require('fs');
var assert = require('assert');
var JSZip = require('jszip');

var excelbuilder = require('..');
var OUTFILE = './test/out/style.xlsx';
var TESTFILE = './test/files/style.xlsx';

var compareWorkbooks = require('./util/compareworkbooks.js')


describe('It generates a simple workbook with styles applied concisely', function () {


  it('generates a ZIP file we can save', function (done) {
    this.timeout(5000)

    var workbook = excelbuilder.createWorkbook()
    var sheet1 = workbook.createSheet('sheet1', 10, 12);

    sheet1.set(1, 1, {
      set: 'Red bold centered  with border',
      font: {
        name: 'Verdana',
        sz: 32,
        color: "FF0022FF",
        bold: true,
        iter: true
      },
      align: 'center',
      fill: {
        type: 'solid',
        fgColor: 'FFFF2200'
      }
    });


    sheet1.set(2, 2, {
      set: Math.PI,
      fill: {
        type: 'solid',
        fgColor: 'FF0022FF'
      },
      numberFormat: '0.00%'
    }) // 10=>'0.00%'


    sheet1.set(3, 3, {
      set: '' + Math.PI,
      fill: {
        type: 'solid',
        fgColor: '99BB66'
      }
    })

    // for some reason date formats only work if the fill is set
    sheet1.set(1, 4, new Date('04/01/2009') )
    sheet1.set(1, 5, {
      set: new Date('04/01/2009'),
      fill: {
        type: "solid",
        fgColor: "FFAA000"
      },
      numberFormat:"m/d/yy"
    } )


    // sheet1.autoFilter(true);
    // Save it
    workbook.generate(function (err, zip) {
      if (err) throw err;
      zip.generateAsync({type: "nodebuffer"}).then(function (buffer) {
        if (err) throw err;
        fs.writeFile(OUTFILE, buffer, function (err) {
          if (err) throw err;
          console.log("open \"" + OUTFILE + "\" ")
          compareWorkbooks(TESTFILE, OUTFILE, function (err, result) {
            if (!result) return done (new Error("Results don't match"))
            done(err);
          })
        })
      })
    })
  })
});
