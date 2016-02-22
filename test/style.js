var fs = require('fs');
var assert = require('assert');
var JSZip = require('jszip');

var excelbuilder = require('..');
var OUTFILE = '/tmp/style.xlsx';
var TESTFILE = './test/files/style.xlsx';


function compareWorkbooks(path1, path2) {
  var zip1 = new JSZip(fs.readFileSync(path1));
  var zip2 = new JSZip(fs.readFileSync(path2));

  for (var key in zip1.files) {
    //console.log(key, zip1.file(key).asText().length, zip2.file(key).asText().length)
    assert.equal(zip1.file(key).asText(), zip2.file(key).asText())
  }
}

describe('It generates a simple workbook', function() {


  it ('generates a ZIP file we can save', function(done) {

    var workbook = excelbuilder.createWorkbook()

    // Create a new worksheet with 10 columns and 12 rows
    var sheet1 = workbook.createSheet('sheet1', 10, 12);

    sheet1.set(1,1,'Red bold centered  with border');
    sheet1.set(2,2,Math.PI);
    sheet1.set(3,3,''+Math.PI);
    sheet1.font(1,1,{
      name: 'Verdana',
      sz: 32,
      color:"FF0022FF",
      bold: true,
      iter:true
    })
    sheet1.align(1,1,'center')
    sheet1.fill(1,1,{
      type: 'solid',
      fgColor: 'FFFF2200'
    })
    sheet1.fill(2,2,{
      type: 'solid',
      fgColor: 'FF0022FF'
    })
    sheet1.fill(3,3,{
      type: 'solid',
      fgColor: 'FF22FF00'
    })
    sheet1.autoFilter(true);
    // Save it
    workbook.generate(function (err, zip) {
      if (err) throw err;
      else {
        var buffer = zip.generate({type: "nodebuffer"});
        fs.writeFile(OUTFILE, buffer, function (err) {
          console.log('Test file written to ' + OUTFILE);
          compareWorkbooks(TESTFILE, OUTFILE)
          done(err);
        });
      }
    });
  })

});

