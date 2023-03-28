var assert = require('assert');
var excelbuilder = require('..');
var JSZip = require('jszip');

describe('jszip', function () {
  it('has generateNodeStream', function (done) {
    var zip = new JSZip();
    zip.generateNodeStream();
    done()
  })
})

describe('msexcel', function () {
  it('has generateNodeStream', function (done) {
    // Create a new workbook file in current working-path
    var path = "/tmp";
    var name = "sample.xlsx";
    var workbook = excelbuilder.createWorkbook(path,name)

    // Create a new worksheet with 10 columns and 12 rows
    var sheet1 = workbook.createSheet('sheet1', 10, 12);

    // Fill some data
    sheet1.set(1, 1, 'I am title');
    for (var i = 2; i < 5; i++)
      sheet1.set(i, 1, 'test' + i);


    workbook.generate(function (err, jszip) {
      if (err)  throw err;
      var stream = jszip.generateNodeStream();
      done();
    });
  })
});
