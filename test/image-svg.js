var fs = require('fs');
var assert = require('assert');
var JSZip = require('jszip');

function requireUncached(module) {
  delete require.cache[require.resolve(module)];
  return require(module);
}

const excelbuilder = requireUncached('..');
// var OUTFILE = './test/out/image-png.xlsx';
var OUTFILE = './lab/image-svg/image-svg.xlsx';
var TESTFILE = './test/files/image-png.xlsx';
var compareWorkbooks = require('./util/compareworkbooks.js')


describe('It applies autofilter', function () {


  it('generates a ZIP file we can save', function (done) {
    this.timeout(20000)
    var workbook = excelbuilder.createWorkbook()

    // Create a new worksheet with 10 columns and 12 rows
    var sheet = workbook.createSheet('TEST', 10, 12);
    sheet.addImage({
      range: 'B3:G6',
        base64: 'PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiID8+CjxzdmcgdmVyc2lvbj0iMS4xIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHhtbG5zOnhsaW5rPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hsaW5rIiB3aWR0aD0iMzAwIiBoZWlnaHQ9IjMwMCIgdmlld0JveD0iMCAwIDMwMCAzMDAiPgoJPGcgaWQ9IndpZGdldCIgdHJhbnNmb3JtPSJtYXRyaXgoMSwwLDAsMSwxNTAsMTUwKSI+CgkJPHBhdGggZD0iTS0xMzMuMTU0LDEzMS4yNjQgTC02Ni41MDQ0LDgyLjU0MDggTC0wLjE2MTUyMywxMzEuMjY0IFoiIGZpbGw9IiMxYWI2OTEiIGZpbGwtb3BhY2l0eT0iMSIgLz4KCQk8cGF0aCBkPSJNMS44OTA1MywxMzEuMjYzIEwxLjg5MDUzLDEzMS4yNjMgQy03MC42NDk2LDEzMS4yNjMgLTEyOS4zNzMsNzIuNTM5NiAtMTI5LjM3MywtMC4wMDA1NTk1ODggTC0xMjkuMzczLC0wLjAwMDU1OTU4OCBDLTEyOS4zNzMsLTcyLjU0MDcgLTcwLjY0OTYsLTEzMS4yNjQgMS44OTA1MywtMTMxLjI2NCBMMS44OTA1MywtMTMxLjI2NCBDNzQuNDMwNywtMTMxLjI2NCAxMzMuMTU0LC03Mi41NDA3IDEzMy4xNTQsLTAuMDAwNTU5NTg4IEwxMzMuMTU0LC0wLjAwMDU1OTU4OCBDMTMzLjE1NCw3Mi41Mzk2IDc0LjQzMDcsMTMxLjI2MyAxLjg5MDUzLDEzMS4yNjMgWiIgZmlsbD0iIzFhYjY5MSIgZmlsbC1vcGFjaXR5PSIxIiAvPgoJCTxwYXRoIGQ9Ik0tNC4xOTMzNywzOS4zMzAzIEwtMjQuMDEwOSw1OS4xNDc4IEwtNzMuNTgzMiw5LjU3NTQ4IEwtNTMuNzY1NywtMTAuMjQyIEwtNC4xOTMzNywzOS4zMzAzIEwtNC4xOTMzNywzOS4zMzAzIFoiIGZpbGw9IiNmZmZmZmYiIGZpbGwtb3BhY2l0eT0iMSIgLz4KCQk8cGF0aCBkPSJNLTIzLjM1ODIsNTkuNzk5NyBMLTQzLjE3NTcsMzkuOTgyMiBMNTYuMjU1NSwtNTkuNDQ5MSBMNzYuMDczLC0zOS42MzE2IEwtMjMuMzU4Miw1OS43OTk3IFoiIGZpbGw9IiNmZmZmZmYiIGZpbGwtb3BhY2l0eT0iMSIgLz4KCTwvZz4KPC9zdmc+',
        extension: 'svg',
      options: {stretch: true}
    })
    sheet.split(2,7)

    workbook.generate(function (err, zip) {
      if (err) throw err;

      else {
        var buffer = zip.generateAsync({type: "nodebuffer"}).then(function (buffer) {
          fs.writeFile(OUTFILE, buffer, function (err) {
            console.log('open \"' + OUTFILE + "\"");
            // compareWorkbooks(TESTFILE, OUTFILE, function (err, result) {
            //   if (!result) return done(new Error("Results don't match"))
            //   assert(result)
              done();
            // })
          })
        })
      }
    })
  })
})
