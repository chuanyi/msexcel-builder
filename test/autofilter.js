var fs = require('fs');
var assert = require('assert');
var JSZip = require('jszip');

var excelbuilder = require('..');
var OUTFILE = './test/out/autofilter.xlsx';
var TESTFILE = './test/files/autofilter.xlsx';
var compareWorkbooks = require('./util/compareworkbooks.js')


describe('It applies autofilter', function () {


  it('generates a ZIP file we can save', function (done) {

    var workbook = excelbuilder.createWorkbook()

    // Create a new worksheet with 10 columns and 12 rows
    var COLS = colors.length;
    var ROWS = 32
    var sheet1 = workbook.createSheet('NEURO RAD', COLS, ROWS + 1);

    for (var c = 0; c < COLS; c++) {
      sheet1.set(c + 1, 1, colors[c]);
    }

    for (var c = 0; c < COLS; c++) {
      for (var r = 0; r < ROWS; r++) {
        sheet1.set(c + 1, r + 2, '' + r * c);
      }
    }

    sheet1.autoFilter(true);

    // Create a new worksheet with 10 columns and 12 rows
    var sheet2 = workbook.createSheet('NEURO ONC', 10, 12);
    var colNames = 'ALPHA,BRAVO,CHARLIE,DELTA,ECHO,FOXTROT,GOLF,HOTEL,INDIA'.split(',');

    for (var c = 0; c < 10; c++) {
      sheet2.set(c + 1, 1, colNames[c]);
    }

    for (var c = 0; c < 10; c++) {
      for (var r = 0; r < 11; r++) {
        sheet2.set(c + 1, r + 2, r * c);
      }
    }

    sheet2.autoFilter('A1:E12');

    workbook.generate(function (err, zip) {
      if (err) throw err;
      else {
        var buffer = zip.generateAsync({type: "nodebuffer"}).then(function (buffer) {
          fs.writeFile(OUTFILE, buffer, function (err) {
            console.log('open \"' + OUTFILE + "\"");
            compareWorkbooks(TESTFILE, OUTFILE, function (err, result) {
              if (!result) return done(new Error("Results don't match"))
              //assert(result)
              done();
            })

          });
        })
      }
    });
  })
});

var colors = ["AliceBlue", "AntiqueWhite", "Aqua", "Aquamarine", "Azure", "Beige", "Bisque", "Black", "BlanchedAlmond", "Blue", "BlueViolet", "Brown", "BurlyWood", "CadetBlue", "Chartreuse", "Chocolate", "Coral", "CornflowerBlue", "Cornsilk", "Crimson", "Cyan", "DarkBlue", "DarkCyan", "DarkGoldenRod", "DarkGray", "DarkGrey", "DarkGreen", "DarkKhaki", "DarkMagenta", "DarkOliveGreen", "DarkOrange", "DarkOrchid", "DarkRed", "DarkSalmon", "DarkSeaGreen", "DarkSlateBlue", "DarkSlateGray", "DarkSlateGrey", "DarkTurquoise", "DarkViolet", "DeepPink", "DeepSkyBlue", "DimGray", "DimGrey", "DodgerBlue", "FireBrick", "FloralWhite", "ForestGreen", "Fuchsia", "Gainsboro", "GhostWhite", "Gold", "GoldenRod", "Gray", "Grey", "Green", "GreenYellow", "HoneyDew", "HotPink", "IndianRed", "Indigo", "Ivory", "Khaki", "Lavender", "LavenderBlush", "LawnGreen", "LemonChiffon", "LightBlue", "LightCoral", "LightCyan", "LightGoldenRodYellow", "LightGray", "LightGrey", "LightGreen", "LightPink", "LightSalmon", "LightSeaGreen", "LightSkyBlue", "LightSlateGray", "LightSlateGrey", "LightSteelBlue", "LightYellow", "Lime", "LimeGreen", "Linen", "Magenta", "Maroon", "MediumAquaMarine", "MediumBlue", "MediumOrchid", "MediumPurple", "MediumSeaGreen", "MediumSlateBlue", "MediumSpringGreen", "MediumTurquoise", "MediumVioletRed", "MidnightBlue", "MintCream", "MistyRose", "Moccasin", "NavajoWhite", "Navy", "OldLace", "Olive", "OliveDrab", "Orange", "OrangeRed", "Orchid", "PaleGoldenRod", "PaleGreen", "PaleTurquoise", "PaleVioletRed", "PapayaWhip", "PeachPuff", "Peru", "Pink", "Plum", "PowderBlue", "Purple", "RebeccaPurple", "Red", "RosyBrown", "RoyalBlue", "SaddleBrown", "Salmon", "SandyBrown", "SeaGreen", "SeaShell", "Sienna", "Silver", "SkyBlue", "SlateBlue", "SlateGray", "SlateGrey", "Snow", "SpringGreen", "SteelBlue", "Tan", "Teal", "Thistle", "Tomato", "Turquoise", "Violet", "Wheat", "White", "WhiteSmoke", "Yellow", "YellowGreen"]