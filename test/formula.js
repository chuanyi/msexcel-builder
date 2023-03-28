const fs = require('fs');
const assert = require('assert');
const JSZip = require('jszip');

function requireUncached(module) {
  delete require.cache[require.resolve(module)];
  return require(module);
}

const excelbuilder = requireUncached('..');
const OUTFILE = './test/out/formula.xlsx';
const TESTFILE = './test/files/formula.xlsx';
const compareWorkbooks = require('./util/compareworkbooks.js')


describe('It applies autofilter', function () {


  it('generates worksheet with formulas', function (done) {

    const workbook = excelbuilder.createWorkbook()

    // Create a new worksheet with 10 columns and 12 rows
    const cols = 26;
    const rows = 52;
    const sheet = workbook.createSheet('TEST', cols, rows);
    const colNames = 'ALPHA,BRAVO,CHARLIE,DELTA,ECHO,FOXTROT,GOLF,HOTEL,INDIA,JULIETT,KILO,LIMA,MIKE,NOVEMBER,OSCAR,PAPA,QUEBEC,ROMEO,SIERRA,TANGO,UNIFORM,VICTOR,WHISKEY,X RAY,YANKEE,ZULU'.split(',');

    for (var c = 0; c < cols; c++) {
      sheet.set(c + 1, 1, c);
    }

    for (var r = 0; r < (rows - 1); r++) {
      sheet.set( 1, r + 1, r);
    }

    for (var c = 0; c < (cols - 1); c++) {
      for (var r = 0; r < (rows - 2); r++) {
        sheet.formula(c + 2, r + 2, `${('ABCDEFGHIJKLMNOPQRSTUVWXYZ')[c+1]}1 * A${r+2}`);
      }
    }

    for (var c = 0; c < (cols - 1); c++) {
        sheet.formula(c + 2, rows, `SUM(${('ABCDEFGHIJKLMNOPQRSTUVWXYZ')[c+1]}2:${('ABCDEFGHIJKLMNOPQRSTUVWXYZ')[c+1]}${rows-1}) - ${('ABCDEFGHIJKLMNOPQRSTUVWXYZ')[c]}${rows}`);
    }


    sheet.split(1,1)

    workbook.generate(function (err, zip) {
      if (err) throw err;
      else {
        var buffer = zip.generateAsync({type: "nodebuffer"}).then(function (buffer) {
          fs.writeFile(OUTFILE, buffer, function (err) {
            console.log('open \"' + OUTFILE + "\"");
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
