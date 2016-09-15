var JSZip = require('jszip');
var fs = require('fs')
var async = require('async');



module.exports = function (path1, path2, callback) {
  new JSZip.loadAsync(fs.readFileSync(path1)).then(function (zip1) {
    new JSZip.loadAsync(fs.readFileSync(path2)).then(function (zip2) {

      function compareOne(key, cb) {
        zip1.file(key).async("string").then(function(text1) {
          zip2.file(key).async("string").then(function(text2) {
            return cb(null, text1==text2);
          })
        })
      }

      async.map(Object.keys(zip1.files), compareOne, function(err, results) {
        if (err) return callback(err)
        for (var i=0; i<results.length; i++) {
          if (!results[i]) return callback(null, false)
        }
        return callback(null, true)
      })
    })
  })
}