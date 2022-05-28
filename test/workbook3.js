var fs = require('fs');
var assert = require('assert');
var JSZip = require('jszip');
var path = require('path')
var util = require('util')
var compareWorkbooks = require('./util/compareworkbooks.js')

function requireUncached(module) {
  delete require.cache[require.resolve(module)];
  return require(module);
}

const excelbuilder = requireUncached('..');
describe('It sets cells ', function () {
  it('Generates workbook from JSON', async function () {
    var OUTFILE = './test/out/workbook3.xlsx';
    var TESTFILE = './test/files/workbook3.xlsx'
    var workbook = excelbuilder.createWorkbook()
    workbook.set(pojo)
    var zip = await workbook.generate()
    var buffer = await zip.generateAsync({type: "nodebuffer"})
    fs.writeFileSync(OUTFILE, buffer)
    console.log("open \"" + OUTFILE + "\"")
    var result = await compareWorkbooks(TESTFILE, OUTFILE)
    if (!result) throw new Error(["Results don't match #1", TESTFILE, OUTFILE].join(":"))
    else return true
  })
})

var pojo = {}


