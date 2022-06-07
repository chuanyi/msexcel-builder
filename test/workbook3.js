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
    return done()
    var result = await compareWorkbooks(TESTFILE, OUTFILE)
    if (!result) throw new Error(["Results don't match #1", TESTFILE, OUTFILE].join(":"))
    else return true
  })
})

var pojo = {
  "indexed": {
    "Intro": true,
    "Q1": true
  },
  "worksheets": [
    {
      "name": "Intro",
      "cells": [
        [],
        [],
        [
          {
            "set": "protobi",
            "font": {
              "sz": 36,
              "bold": true,
              "color": {
                "rgb": "FF3399CC"
              },
              "iter": "-",
              "name": "Calibri",
              "scheme": "minor",
              "family": "2",
              "underline": "-",
              "strike": "-",
              "outline": "-",
              "shadow": "-"
            }
          }
        ],
        [],
        [
          {
            "set": "Project: ",
            "font": {
              "bold": true,
              "iter": "-",
              "sz": "11",
              "color": "-",
              "name": "Calibri",
              "scheme": "minor",
              "family": "2",
              "underline": "-",
              "strike": "-",
              "outline": "-",
              "shadow": "-"
            }
          },
          {
            "set": "Example Pharma",
            "width": 32
          }
        ],
        [
          {
            "set": "Generated: ",
            "font": {
              "bold": true,
              "iter": "-",
              "sz": "11",
              "color": "-",
              "name": "Calibri",
              "scheme": "minor",
              "family": "2",
              "underline": "-",
              "strike": "-",
              "outline": "-",
              "shadow": "-"
            }
          },
          {
            "set": "07-06-2022 12:29 pm"
          }
        ],
        [
          {
            "set": "URL: ",
            "font": {
              "bold": true,
              "iter": "-",
              "sz": "11",
              "color": "-",
              "name": "Calibri",
              "scheme": "minor",
              "family": "2",
              "underline": "-",
              "strike": "-",
              "outline": "-",
              "shadow": "-"
            }
          },
          {
            "set": "http://localhost:3000/api/v3/datasets/example/app.html#filter/%7B%22global%22%3A%7B%7D%7D/options/%7B%22showMissing%22%3Atrue%2C%22showPercent%22%3Atrue%2C%22showFormat%22%3Atrue%2C%22showWeighted%22%3Atrue%7D",
            "l": {
              "target": "http://localhost:3000/api/v3/datasets/example/app.html#filter/%7B%22global%22%3A%7B%7D%7D/options/%7B%22showMissing%22%3Atrue%2C%22showPercent%22%3Atrue%2C%22showFormat%22%3Atrue%2C%22showWeighted%22%3Atrue%7D"
            }
          }
        ],
        [],
        [
          {
            "set": "Current scenario: ",
            "font": {
              "bold": true,
              "iter": "-",
              "sz": "11",
              "color": "-",
              "name": "Calibri",
              "scheme": "minor",
              "family": "2",
              "underline": "-",
              "strike": "-",
              "outline": "-",
              "shadow": "-"
            }
          }
        ],
        [
          {
            "set": "N: ",
            "font": {
              "bold": "-",
              "iter": "-",
              "sz": "11",
              "color": "-",
              "name": "Calibri",
              "scheme": "minor",
              "family": "2",
              "underline": "-",
              "strike": "-",
              "outline": "-",
              "shadow": "-"
            }
          },
          100
        ],
        [
          "Filters:",
          null
        ],
        [],
        [
          {
            "set": "Baseline: ",
            "font": {
              "bold": true,
              "iter": "-",
              "sz": "11",
              "color": "-",
              "name": "Calibri",
              "scheme": "minor",
              "family": "2",
              "underline": "-",
              "strike": "-",
              "outline": "-",
              "shadow": "-"
            }
          }
        ],
        [
          {
            "set": "N: ",
            "font": {
              "bold": "-",
              "iter": "-",
              "sz": "11",
              "color": "-",
              "name": "Calibri",
              "scheme": "minor",
              "family": "2",
              "underline": "-",
              "strike": "-",
              "outline": "-",
              "shadow": "-"
            }
          },
          100
        ],
        [
          "Filters:",
          null
        ],
        [],
        [
          {
            "set": "Contents: ",
            "font": {
              "bold": true,
              "iter": "-",
              "sz": "11",
              "color": "-",
              "name": "Calibri",
              "scheme": "minor",
              "family": "2",
              "underline": "-",
              "strike": "-",
              "outline": "-",
              "shadow": "-"
            }
          }
        ],
        [
          "",
          {
            "set": "Element ",
            "font": {
              "bold": true,
              "iter": "-",
              "sz": "11",
              "color": "-",
              "name": "Calibri",
              "scheme": "minor",
              "family": "2",
              "underline": "-",
              "strike": "-",
              "outline": "-",
              "shadow": "-"
            }
          },
          {
            "set": "Title ",
            "font": {
              "bold": true,
              "iter": "-",
              "sz": "11",
              "color": "-",
              "name": "Calibri",
              "scheme": "minor",
              "family": "2",
              "underline": "-",
              "strike": "-",
              "outline": "-",
              "shadow": "-"
            }
          }
        ],
        [
          "",
          {
            "set": "Q1"
          },
          {
            "set": "Thinking about your Condition X patients, how many of those are currently receiving a GA (gamma antagonist)?"
          }
        ],
        [],
        [
          {
            "set": "(c) protobi 2022 All rights reserved.",
            "font": {
              "sz": 12,
              "bold": "-",
              "color": {
                "rgb": "FF3399CC"
              },
              "iter": "-",
              "name": "Calibri",
              "scheme": "minor",
              "family": "2",
              "underline": "-",
              "strike": "-",
              "outline": "-",
              "shadow": "-"
            }
          }
        ]
      ],
      "options": {
        "sheetViews": {
          "showGridLines": "0"
        }
      }
    },
    {
      "name": "Q1",
      "cells": [
        [],
        [],
        [],
        [
          {
            "set": "Q1",
            "font": {
              "bold": true,
              "iter": "-",
              "sz": "11",
              "color": "-",
              "name": "Calibri",
              "scheme": "minor",
              "family": "2",
              "underline": "-",
              "strike": "-",
              "outline": "-",
              "shadow": "-"
            },
            "width": 32
          },
          "Thinking about your Condition X patients, how many of those are currently receiving a GA (gamma antagonist)?"
        ],
        [
          {
            "set": "Value",
            "font": {
              "bold": true,
              "iter": "-",
              "sz": "11",
              "color": "-",
              "name": "Calibri",
              "scheme": "minor",
              "family": "2",
              "underline": "-",
              "strike": "-",
              "outline": "-",
              "shadow": "-"
            }
          },
          {
            "set": "current",
            "font": {
              "bold": true,
              "iter": "-",
              "sz": "11",
              "color": "-",
              "name": "Calibri",
              "scheme": "minor",
              "family": "2",
              "underline": "-",
              "strike": "-",
              "outline": "-",
              "shadow": "-"
            }
          },
          null,
          {
            "set": "Value",
            "font": {
              "bold": true,
              "iter": "-",
              "sz": "11",
              "color": "-",
              "name": "Calibri",
              "scheme": "minor",
              "family": "2",
              "underline": "-",
              "strike": "-",
              "outline": "-",
              "shadow": "-"
            }
          },
          {
            "set": "current",
            "font": {
              "bold": true,
              "iter": "-",
              "sz": "11",
              "color": "-",
              "name": "Calibri",
              "scheme": "minor",
              "family": "2",
              "underline": "-",
              "strike": "-",
              "outline": "-",
              "shadow": "-"
            }
          }
        ],
        [
          {
            "set": "0"
          },
          {
            "set": 0.05,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          {
            "set": "0"
          },
          {
            "set": 5,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          }
        ],
        [
          {
            "set": "1 to 15"
          },
          {
            "set": 0.7,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          {
            "set": "1 to 15"
          },
          {
            "set": 70,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          }
        ],
        [
          {
            "set": "16 to 30"
          },
          {
            "set": 0.16,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          {
            "set": "16 to 30"
          },
          {
            "set": 16,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          }
        ],
        [
          {
            "set": "31 to 45"
          },
          {
            "set": 0.02,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          {
            "set": "31 to 45"
          },
          {
            "set": 2,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          }
        ],
        [
          {
            "set": "46 to 60"
          },
          {
            "set": 0.04,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          {
            "set": "46 to 60"
          },
          {
            "set": 4,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          }
        ],
        [
          {
            "set": "61 to 75"
          },
          {
            "set": 0.01,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          {
            "set": "61 to 75"
          },
          {
            "set": 1,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          }
        ],
        [
          {
            "set": "76 to 90"
          },
          {
            "set": 0.01,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          {
            "set": "76 to 90"
          },
          {
            "set": 1,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          }
        ],
        [
          {
            "set": "226 to 240"
          },
          {
            "set": 0.01,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          {
            "set": "226 to 240"
          },
          {
            "set": 1,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          }
        ],
        [
          "N",
          {
            "set": 100,
            "font": {
              "italic": true,
              "color": {
                "theme": 3
              },
              "bold": "-",
              "iter": "-",
              "sz": "11",
              "name": "Calibri",
              "scheme": "minor",
              "family": "2",
              "underline": "-",
              "strike": "-",
              "outline": "-",
              "shadow": "-"
            },
            "numberFormat": "\"\"0.0;\"\"-0.0;\"\"\\—;@"
          },
          null,
          "N",
          {
            "set": 100,
            "font": {
              "italic": true,
              "color": {
                "theme": 3
              },
              "bold": "-",
              "iter": "-",
              "sz": "11",
              "name": "Calibri",
              "scheme": "minor",
              "family": "2",
              "underline": "-",
              "strike": "-",
              "outline": "-",
              "shadow": "-"
            },
            "numberFormat": "\"\"0.0;\"\"-0.0;\"\"\\—;@"
          }
        ]
      ],
      "options": {
        "sheetViews": {
          "showGridLines": "0"
        }
      }
    }
  ]
}


