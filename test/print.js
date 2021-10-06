var fs = require('fs');
var assert = require('assert');
var JSZip = require('jszip');

function requireUncached(module) {
  delete require.cache[require.resolve(module)];
  return require(module);
}

const excelbuilder = requireUncached('..');
var OUTFILE = './lab/print/print.xlsx';
var TESTFILE = './test/files/print.xlsx';
var compareWorkbooks = require('./util/compareworkbooks.js')


describe('It applies autofilter', function () {


  it('generates a ZIP file we can save', function (done) {

    var workbook = excelbuilder.createWorkbook()

    // Create a new worksheet with 10 columns and 12 rows
    var sheet = workbook.createSheet('TEST', 110, 210);
    var colNames = 'ALPHA,BRAVO,CHARLIE,DELTA,ECHO,FOXTROT,GOLF,HOTEL,INDIA'.split(',');

    for (var c = 0; c < 100; c++) {
      sheet.set(c + 1, 1, colNames[c]);
      for (var r = 0; r < 200; r++) {
        sheet.set(c + 1, r + 2, r * c);
      }
    }

    sheet.autoFilter('A1:E12');
    sheet.sheetViews({
      showGridLines: "0",
    })
    sheet.pageMargins({
      left: 0.25,
      right: 0.25,
      top: 0.5,
      bottom: 0.5,
      header: 0.3,
      footer: 0.3

    })
    sheet.pageSetup({
      scale: "50",
    })
    sheet.printBreakColumns([15,30,45,60,75,90,105,120,135,150])
    sheet.printBreakRows([90,180,270])

    sheet.printRepeatRows(1,3)
    sheet.printRepeatColumns(1,2)

    sheet.split(1,1)

    workbook.generate(function (err, zip) {
      if (err) throw err;
      else {
        var buffer = zip.generateAsync({type: "nodebuffer"}).then(function (buffer) {
          fs.writeFile(OUTFILE, buffer, function (err) {
            console.log('open \"' + OUTFILE + "\"");
            // done();
            compareWorkbooks(TESTFILE, OUTFILE, function (err, result) {
              if (!result) return done(new Error("Results don't match"))
              assert(result)
              done();
            })
          })
        })
      }
    })
  })
})

var colors = ["AliceBlue", "AntiqueWhite", "Aqua", "Aquamarine", "Azure", "Beige", "Bisque", "Black", "BlanchedAlmond", "Blue", "BlueViolet", "Brown", "BurlyWood", "CadetBlue", "Chartreuse", "Chocolate", "Coral", "CornflowerBlue", "Cornsilk", "Crimson", "Cyan", "DarkBlue", "DarkCyan", "DarkGoldenRod", "DarkGray", "DarkGrey", "DarkGreen", "DarkKhaki", "DarkMagenta", "DarkOliveGreen", "DarkOrange", "DarkOrchid", "DarkRed", "DarkSalmon", "DarkSeaGreen", "DarkSlateBlue", "DarkSlateGray", "DarkSlateGrey", "DarkTurquoise", "DarkViolet", "DeepPink", "DeepSkyBlue", "DimGray", "DimGrey", "DodgerBlue", "FireBrick", "FloralWhite", "ForestGreen", "Fuchsia", "Gainsboro", "GhostWhite", "Gold", "GoldenRod", "Gray", "Grey", "Green", "GreenYellow", "HoneyDew", "HotPink", "IndianRed", "Indigo", "Ivory", "Khaki", "Lavender", "LavenderBlush", "LawnGreen", "LemonChiffon", "LightBlue", "LightCoral", "LightCyan", "LightGoldenRodYellow", "LightGray", "LightGrey", "LightGreen", "LightPink", "LightSalmon", "LightSeaGreen", "LightSkyBlue", "LightSlateGray", "LightSlateGrey", "LightSteelBlue", "LightYellow", "Lime", "LimeGreen", "Linen", "Magenta", "Maroon", "MediumAquaMarine", "MediumBlue", "MediumOrchid", "MediumPurple", "MediumSeaGreen", "MediumSlateBlue", "MediumSpringGreen", "MediumTurquoise", "MediumVioletRed", "MidnightBlue", "MintCream", "MistyRose", "Moccasin", "NavajoWhite", "Navy", "OldLace", "Olive", "OliveDrab", "Orange", "OrangeRed", "Orchid", "PaleGoldenRod", "PaleGreen", "PaleTurquoise", "PaleVioletRed", "PapayaWhip", "PeachPuff", "Peru", "Pink", "Plum", "PowderBlue", "Purple", "RebeccaPurple", "Red", "RosyBrown", "RoyalBlue", "SaddleBrown", "Salmon", "SandyBrown", "SeaGreen", "SeaShell", "Sienna", "Silver", "SkyBlue", "SlateBlue", "SlateGray", "SlateGrey", "Snow", "SpringGreen", "SteelBlue", "Tan", "Teal", "Thistle", "Tomato", "Turquoise", "Violet", "Wheat", "White", "WhiteSmoke", "Yellow", "YellowGreen"]
