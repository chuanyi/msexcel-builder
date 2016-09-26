var JSZip = require('jszip');
var fs = require('fs')
var async = require('async');


module.exports = function (path1, path2, callback) {
  new JSZip.loadAsync(fs.readFileSync(path1)).then(function (zip1) {
    new JSZip.loadAsync(fs.readFileSync(path2)).then(function (zip2) {

      function compareOne(key, cb) {

        if ([".xml", "rels"].indexOf(key.substr(-4)) >= 0) {
          zip1.file(key).async("string").then(function (text1) {

            zip2.file(key).async("string").then(function (text2) {

              return cb(null, text1 == text2);
            }).catch(function (err) {
              return cb(err);
            })
          }).catch(function (err) {
            return cb(err);
          })
        }
        else {
          return cb(null, true)
        }
      }

      var fileKeys = Object.keys(zip1.files);
      async.map(fileKeys, compareOne, function (err, results) {
        if (err) return callback(err)
        
        else {

          for (var i = 0; i < results.length; i++) {
            if (!results[i]) {
              return callback(null, false)
            }
          }
          return callback(null, true)
        }
      })
    })
  })
}