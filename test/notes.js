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


    sheet1.set(1, 1,);
    for (var i = 2; i < 6; i++) {
      sheet1.set(i, 1, 'test' + i);
      sheet1.set(i, 2, i / 2);

    }

    var lyrics = [
      "Somebody once told me the world is gonna roll me",
      "I ain't the sharpest tool in the shed",
      "She was looking kind of dumb with her finger and her thumb",
      "In the shape of an \"L\" on her forehead",
      "",
      "Well the years start coming and they don't stop coming",
      "Fed to the rules and I hit the ground running",
      "Didn't make sense not to live for fun",
      "Your brain gets smart but your head gets dumb",
      "",
      "So much to do, so much to see",
      "So what's wrong with taking the back streets?",
      "You'll never know if you don't go",
      "You'll never shine if you don't glow",
      "",
      "Hey now, you're an all-star, get your game on, go play",
      "Hey now, you're a rock star, get the show on, get paid",
      "And all that glitters is gold",
      "Only shooting stars break the mold",
      "",
      "-- Smashmouth",
    ]
    sheet1.note(2, 2, lyrics)
    sheet1.note(3, 2, {text: `Berry`, props: {bold: true}})
    sheet1.note(4, 2, `Cherry`)
    sheet1.note(5, 2, `Date`)
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


//<rPr><b/><sz val="11"/><color rgb="FF000000"/><rFont val="Calibri"/><family val="2"/><charset val="134"/></rPr><t xml:space="preserve">
//, fontSize: 14, fontFamily: "Courier New"
