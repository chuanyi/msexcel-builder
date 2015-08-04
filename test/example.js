var fs = require('fs');
var assert = require('assert');
var JSZip = require('jszip');

var excelbuilder = require('..');

function compareWorkbooks(path1, path2) {
  var zip1 = new JSZip(fs.readFileSync(path1));
  var zip2 = new JSZip(fs.readFileSync(path2));

  for (var key in zip1.files) {
    //console.log(key, zip1.file(key).asText().length, zip2.file(key).asText().length)
    assert.equal(zip1.file(key).asText(), zip2.file(key).asText())
  }
}

describe('It generates a simple workbook', function() {

  it ('has a vestigial cancel method for backward compatibility', function() {
    var workbook = excelbuilder.createWorkbook()
    workbook.cancel()
  })

  it ('generates a ZIP file we can save', function(done) {

    var workbook = excelbuilder.createWorkbook()

// Create a new worksheet with 10 columns and 12 rows
    var sheet1 = workbook.createSheet('sheet1', 10, 12);

// Fill some data
    sheet1.set(1, 1, 'I am title');
    for (var i = 2; i < 5; i++)
      sheet1.set(i, 1, 'test' + i);

// Save it
    workbook.generate(function (err, zip) {
      if (err) throw err;
      else {
        var buffer = zip.generate({type: "nodebuffer"});
        var OUTFILE = '/tmp/example.xlsx';
        fs.writeFile(OUTFILE, buffer, function (err) {
          console.log('Test file written to ' + OUTFILE);
          compareWorkbooks('./test/files/example.xlsx', OUTFILE)
          done(err);
        });
      }
    });
  })

  it ('Supports the prior constructor syntax', function(done) {
    var PATH = '/tmp';
    var FILENAME = 'example2.xlsx';
    var workbook = excelbuilder.createWorkbook(PATH, FILENAME);
    var sheet1 = workbook.createSheet('sheet1', 10, 12);

    sheet1.set(1, 1, 'I am title');
    for (var i = 2; i < 5; i++)
      sheet1.set(i, 1, 'test' + i);

    workbook.save(function (err) {
      if (err) throw err;
      else {
        var OUTFILE = PATH + "/" + FILENAME;
        console.log('Test file written to ' + OUTFILE);
        done()
      }
    });
  })
});

