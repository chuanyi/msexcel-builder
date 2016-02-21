var fs = require('fs');
var assert = require('assert');
var JSZip = require('jszip');

var excelbuilder = require('..');
var OUTFILE = '/tmp/autofilter.xlsx';
var TESTFILE = './test/files/autofilter.xlsx';

function compareWorkbooks(path1, path2) {
  var zip1 = new JSZip(fs.readFileSync(path1));
  var zip2 = new JSZip(fs.readFileSync(path2));

  for (var key in zip1.files) {
    //console.log(key, zip1.file(key).asText().length, zip2.file(key).asText().length)
    assert.equal(zip1.file(key).asText(), zip2.file(key).asText())
  }
}

describe('It applies autofilter', function() {


  it ('generates a ZIP file we can save', function(done) {

    var workbook = excelbuilder.createWorkbook()

    // Create a new worksheet with 10 columns and 12 rows
    var sheet1 = workbook.createSheet('sheet1', 10, 12);
    var colNames = 'ALPHA,BRAVO,CHARLIE,DELTA,ECHO,FOXTROT,GOLF,HOTEL,INDIA'.split(',');

    for (var c=0; c<10; c++) {
      sheet1.set(c+1,1, colNames[c]);
    }

    for (var c=0; c<10; c++) {
      for (var r=0; r<11; r++) {
        sheet1.set(c+1,r+2, ''+r*c);
      }
    }

    sheet1.autoFilter(true);

    // Create a new worksheet with 10 columns and 12 rows
    var sheet2 = workbook.createSheet('sheet2', 10, 12);
    var colNames = 'ALPHA,BRAVO,CHARLIE,DELTA,ECHO,FOXTROT,GOLF,HOTEL,INDIA'.split(',');

    for (var c=0; c<10; c++) {
      sheet2.set(c+1,1, colNames[c]);
    }

    for (var c=0; c<10; c++) {
      for (var r=0; r<11; r++) {
        sheet2.set(c+1,r+2,r*c);
      }
    }

    sheet2.autoFilter('A1:E12');

    workbook.generate(function (err, zip) {
      if (err) throw err;
      else {
        var buffer = zip.generate({type: "nodebuffer"});
        fs.writeFile(OUTFILE, buffer, function (err) {
          console.log('open ' + OUTFILE);
//          compareWorkbooks(TESTFILE, OUTFILE)
          done(err);
        });
      }
    });
  })

});

