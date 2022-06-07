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

var pojo = {
  "indexed": {
    "Intro": true,
    "Section 1": true,
    "Section 2": true
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
            "set": "07-06-2022 02:18 pm"
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
            "set": "Section 1"
          },
          {
            "set": "Current prescribing "
          }
        ],
        [
          "",
          {
            "set": "Section 2"
          },
          {
            "set": "Product profile"
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
      "sheetViews": {
        "showGridLines": "0"
      }
    },
    {
      "name": "Section 1",
      "cells": [
        [],
        [],
        [],
        [
          {
            "set": "Section 1",
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
          {
            "set": "Current prescribing "
          }
        ],
        [],
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
        ],
        [],
        [],
        [
          {
            "set": "Q2",
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
          {
            "set": "Thinking about your Condition X patients on a  GA (gamma antagonist), what percent are currently on the following therapies? (Your answers should sum to 100)"
          }
        ],
        [
          {
            "set": "vs. Q1",
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
          {
            "set": "Thinking about your Condition X patients, how many of those are currently receiving a GA (gamma antagonist)?"
          }
        ],
        [
          {
            "set": "vs. S1",
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
          {
            "set": "What is your specialty (select only one answer) "
          }
        ],
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
          {
            "set": "0",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            },
            "width": 12
          },
          null,
          null,
          {
            "set": "1 to 15",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            },
            "width": 12
          },
          null,
          null,
          {
            "set": "16 to 30",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            },
            "width": 12
          },
          null,
          null,
          {
            "set": "31 to 45",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            },
            "width": 12
          },
          null,
          {
            "set": "46 to 60",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            },
            "width": 12
          },
          null,
          null,
          {
            "set": "61 to 75",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            },
            "width": 12
          },
          null,
          {
            "set": "76 to 90",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            },
            "width": 12
          },
          null,
          {
            "set": "226 to 240",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            },
            "width": 12
          }
        ],
        [
          {
            "set": "S1",
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
            "border": {
              "bottom": {
                "style": "thin",
                "color": "888888"
              },
              "left": "-",
              "right": "-",
              "top": "-"
            },
            "width": 32
          },
          {
            "set": "Overall",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "bottom": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-"
            },
            "width": 12
          },
          {
            "set": "General Practitioner",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "bottom": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-"
            },
            "width": 12
          },
          {
            "set": "Practice Nurse",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "bottom": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-"
            },
            "width": 12
          },
          {
            "set": "Overall",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "bottom": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-"
            },
            "width": 12
          },
          {
            "set": "General Practitioner",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "bottom": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-"
            },
            "width": 12
          },
          {
            "set": "Practice Nurse",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "bottom": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-"
            },
            "width": 12
          },
          {
            "set": "Overall",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "bottom": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-"
            },
            "width": 12
          },
          {
            "set": "General Practitioner",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "bottom": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-"
            },
            "width": 12
          },
          {
            "set": "Practice Nurse",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "bottom": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-"
            },
            "width": 12
          },
          {
            "set": "Overall",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "bottom": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-"
            },
            "width": 12
          },
          {
            "set": "Practice Nurse",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "bottom": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-"
            },
            "width": 12
          },
          {
            "set": "Overall",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "bottom": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-"
            },
            "width": 12
          },
          {
            "set": "General Practitioner",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "bottom": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-"
            },
            "width": 12
          },
          {
            "set": "Practice Nurse",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "bottom": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-"
            },
            "width": 12
          },
          {
            "set": "Overall",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "bottom": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-"
            },
            "width": 12
          },
          {
            "set": "General Practitioner",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "bottom": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-"
            },
            "width": 12
          },
          {
            "set": "Overall",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "bottom": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-"
            },
            "width": 12
          },
          {
            "set": "Practice Nurse",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "bottom": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-"
            },
            "width": 12
          },
          {
            "set": "Overall",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "bottom": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-"
            },
            "width": 12
          },
          {
            "set": "Practice Nurse",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "bottom": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-"
            },
            "width": 12
          }
        ],
        [
          {
            "set": "N",
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
          {
            "set": 5,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          },
          {
            "set": 2,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          },
          {
            "set": 3,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          },
          {
            "set": 70,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          },
          {
            "set": 44,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          },
          {
            "set": 26,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          },
          {
            "set": 16,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          },
          {
            "set": 10,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          },
          {
            "set": 6,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          },
          {
            "set": 2,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          },
          {
            "set": 2,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          },
          {
            "set": 4,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          },
          {
            "set": 3,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          },
          {
            "set": 1,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          },
          {
            "set": 1,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          },
          {
            "set": 1,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          },
          {
            "set": 1,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          },
          {
            "set": 1,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          },
          {
            "set": 1,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          },
          {
            "set": 1,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Product A",
          "",
          "",
          "",
          {
            "set": 34.371428571428574,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 39.40909090909091,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 25.846153846153847,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 42.0625,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 36.3,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 51.666666666666664,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 35,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 35,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 42.5,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 40,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 50,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 30,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 30,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          null,
          null,
          {
            "set": 70,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 70,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          }
        ],
        [
          "Product B",
          "",
          "",
          "",
          {
            "set": 26.442857142857143,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 24.09090909090909,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 30.423076923076923,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 24.75,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 34.1,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFCCDDFF",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF2244AA",
              "iter": "-",
              "sz": "11",
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
            "set": 9.166666666666666,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 15,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 15,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 22.5,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 30,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          null,
          {
            "set": 20,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 20,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 99,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 99,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 20,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 20,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          }
        ],
        [
          "Product C",
          "",
          "",
          "",
          {
            "set": 7.685714285714286,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 6.363636363636363,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 9.923076923076923,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 3.1875,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 5.1,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFCCDDFF",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF2244AA",
              "iter": "-",
              "sz": "11",
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
            "set": 5,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 5,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 2.5,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 3.3333333333333335,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          null,
          {
            "set": 10,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 10,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          null,
          null,
          null,
          null
        ],
        [
          "Product D",
          "",
          "",
          "",
          {
            "set": 5.485714285714286,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 5.113636363636363,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 6.115384615384615,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 5.9375,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 7.5,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 3.3333333333333335,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 5,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 5,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 17.5,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 10,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 40,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 20,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 20,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          null,
          null,
          {
            "set": 5,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 5,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          }
        ],
        [
          "Product E",
          "",
          "",
          "",
          {
            "set": 7.728571428571429,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 10.227272727272727,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 3.5,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 7.3125,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 9.7,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 3.3333333333333335,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 10,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 10,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 8.75,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 8.333333333333334,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 10,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 10,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 10,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 1,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 1,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 2,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 2,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          }
        ],
        [
          "Product F",
          "",
          "",
          "",
          {
            "set": 13,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 11.272727272727273,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 15.923076923076923,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 14.5625,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 6.8,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 27.5,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 30,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 30,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 6.25,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 8.333333333333334,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          null,
          {
            "set": 10,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 10,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          null,
          null,
          {
            "set": 3,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 3,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          }
        ],
        [
          "Other",
          "",
          "",
          "",
          {
            "set": 5.285714285714286,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 3.522727272727273,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 8.26923076923077,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 2.1875,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 0.5,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          {
            "set": 5,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          },
          null,
          null,
          null,
          null,
          null,
          null,
          null,
          null,
          null,
          null,
          null
        ],
        [],
        [],
        [
          {
            "set": "Q2",
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
          {
            "set": "Thinking about your Condition X patients on a  GA (gamma antagonist), what percent are currently on the following therapies? (Your answers should sum to 100)"
          }
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
          }
        ],
        [
          {
            "set": "N",
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
        ],
        [
          "Product D",
          {
            "set": 35.98947368421052,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          }
        ],
        [
          "Product B",
          {
            "set": 26.378947368421052,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          }
        ],
        [
          "Product F",
          {
            "set": 6.515789473684211,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          }
        ],
        [
          "Product C",
          {
            "set": 6.147368421052631,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          }
        ],
        [
          "Product A",
          {
            "set": 7.6421052631578945,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          }
        ],
        [
          "Product E",
          {
            "set": 13.063157894736841,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          }
        ],
        [
          "other",
          {
            "set": 4.2631578947368425,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          }
        ],
        [],
        [],
        [],
        [
          {
            "set": "Q3",
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
          "Thinking about your Condition Y patients, how many of those are currently receiving a GA (gamma antagonist)?"
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
            "set": "1 to 15"
          },
          {
            "set": 0.51,
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
            "set": 51,
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
            "set": 0.32,
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
            "set": 32,
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
            "set": 0.07,
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
            "set": 7,
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
            "set": 0.03,
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
            "set": 3,
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
            "set": "61 to 75"
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
            "set": "91 to 105"
          },
          {
            "set": 0.03,
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
            "set": "91 to 105"
          },
          {
            "set": 3,
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
        ],
        [],
        [],
        [
          {
            "set": "Q4v1",
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
          {
            "set": "Thinking about your Condition Y patients on a GA (gamma antagonist), what percent are currently on the following therapies? (Your answers should sum to 100)\n\n"
          }
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
          }
        ],
        [
          {
            "set": "N",
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
        ],
        [
          "Product D",
          {
            "set": 0.36500000000000005,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          }
        ],
        [
          "Product B",
          {
            "set": 0.1857,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          }
        ],
        [
          "Product F",
          {
            "set": 0.08779999999999999,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          }
        ],
        [
          "Product C",
          {
            "set": 0.07449999999999998,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          }
        ],
        [
          "Product A",
          {
            "set": 0.11189999999999997,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          }
        ],
        [
          "Product E",
          {
            "set": 0.1546,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          }
        ],
        [
          "other",
          {
            "set": 0.020499999999999997,
            "numberFormat": "\"\"0.00;\"\"-0.00;\"\"\\—;@"
          }
        ]
      ],
      "sheetViews": {
        "showGridLines": "0"
      }
    },
    {
      "name": "Section 2",
      "cells": [
        [],
        [],
        [],
        [
          {
            "set": "Section 2",
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
          {
            "set": "Product profile"
          }
        ],
        [],
        [],
        [],
        [],
        [
          {
            "set": "Q7",
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
          "What are your initial thoughts about the device you have just reviewed?   "
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
            "set": "good"
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
            "set": "good"
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
            "set": "interesting, should be easy to use"
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
            "set": "interesting, should be easy to use"
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
            "set": "Good. Positive"
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
            "set": "Good. Positive"
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
            "set": "this product can help the elderly patients to be more independent with their therapy"
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
            "set": "this product can help the elderly patients to be more independent with their therapy"
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
            "set": "Good because it reusable.  Could still be difficult with people who have arthritic problems"
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
            "set": "Good because it reusable.  Could still be difficult with people who have arthritic problems"
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
            "set": "Excellent idea"
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
            "set": "Excellent idea"
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
            "set": "interesting concept but difficult to prescribe on repeat and thus facilitate the device change at 6 months"
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
            "set": "interesting concept but difficult to prescribe on repeat and thus facilitate the device change at 6 months"
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
            "set": "Sounds good that it is refillable for 6/12. Like the larger counter and red warning of low doses left The fact it's bigger good for elderly patients not sure younger patients will be so pleased . Don't like the cap over mouthpiece looks very vulnerable to damage."
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
            "set": "Sounds good that it is refillable for 6/12. Like the larger counter and red warning of low doses left The fact it's bigger good for elderly patients not sure younger patients will be so pleased . Don't like the cap over mouthpiece looks very vulnerable to damage."
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
            "set": "Interesting and important development with lesser impact on environment ."
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
            "set": "Interesting and important development with lesser impact on environment ."
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
            "set": "looking forward to it. seems easy enough to do"
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
            "set": "looking forward to it. seems easy enough to do"
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
            "set": "Difficult to load for the elderly"
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
            "set": "Difficult to load for the elderly"
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
            "set": "good, it may save cost"
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
            "set": "good, it may save cost"
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
            "set": "Too complicated for compliance-looks overwheming"
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
            "set": "Too complicated for compliance-looks overwheming"
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
            "set": "Good.  I like the Product B device but it's not great for patients with poor dexterity"
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
            "set": "Good.  I like the Product B device but it's not great for patients with poor dexterity"
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
            "set": "Would appear to be a very user freindly device. Clear doseage indicator and easy to handle."
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
            "set": "Would appear to be a very user freindly device. Clear doseage indicator and easy to handle."
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
            "set": "It seems relatively simple to use and cost effective. Reducing waste"
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
            "set": "It seems relatively simple to use and cost effective. Reducing waste"
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
            "set": "Looks quite a good product"
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
            "set": "Looks quite a good product"
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
            "set": "Lack of guidelines"
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
            "set": "Lack of guidelines"
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
            "set": "reasonable device but still a bit cumbersome to operate especially for very old folk and people with arthritis"
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
            "set": "reasonable device but still a bit cumbersome to operate especially for very old folk and people with arthritis"
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
            "set": "looks like a similar device  I have seen , compact easy dosing, easy to use"
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
            "set": "looks like a similar device  I have seen , compact easy dosing, easy to use"
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
            "set": "hopefully easy to handle, daily use looks fiddly"
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
            "set": "hopefully easy to handle, daily use looks fiddly"
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
            "set": "Reusable a plus. Looks easy to use and handle. Easy to read"
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
            "set": "Reusable a plus. Looks easy to use and handle. Easy to read"
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
            "set": "Looks similar to Product B. The market place is already very crowdwd"
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
            "set": "Looks similar to Product B. The market place is already very crowdwd"
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
            "set": "Easy  to use and refill environmental friendly possible cost effective"
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
            "set": "Easy  to use and refill environmental friendly possible cost effective"
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
            "set": "looks useful and superior to current devices - like that the medicine clears"
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
            "set": "looks useful and superior to current devices - like that the medicine clears"
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
            "set": "Looks easy to use but may be too bulky for some"
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
            "set": "Looks easy to use but may be too bulky for some"
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
            "set": "It looks user friendly. The large size will be popular with most patients except those who travel frequently."
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
            "set": "It looks user friendly. The large size will be popular with most patients except those who travel frequently."
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
            "set": "looks competent,accurate and user feiendly"
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
            "set": "looks competent,accurate and user feiendly"
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
            "set": "innovative simple friendly"
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
            "set": "innovative simple friendly"
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
            "set": "Monthly refill. Change 6 months. Large dose indicator compared to Product B."
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
            "set": "Monthly refill. Change 6 months. Large dose indicator compared to Product B."
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
            "set": "Looks very user friendly and I am happy it can be reused"
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
            "set": "Looks very user friendly and I am happy it can be reused"
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
            "set": "Looks easier load than resonate. Less wastage"
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
            "set": "Looks easier load than resonate. Less wastage"
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
            "set": "Good for all ages, especially the elderly. Good to know when device nearly empty."
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
            "set": "Good for all ages, especially the elderly. Good to know when device nearly empty."
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
            "set": "Looks novel"
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
            "set": "Looks novel"
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
            "set": "Looks easy to use - bigger device, easier to handle and manipulate, easier to read dose counter. Looks aesthetically pleasing."
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
            "set": "Looks easy to use - bigger device, easier to handle and manipulate, easier to read dose counter. Looks aesthetically pleasing."
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
            "set": "looks a bit complicated for patient, the instructions are very overwhelming"
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
            "set": "looks a bit complicated for patient, the instructions are very overwhelming"
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
            "set": "looks good but bulky. nothing inspiring. low tech"
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
            "set": "looks good but bulky. nothing inspiring. low tech"
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
            "set": "similar to others"
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
            "set": "similar to others"
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
            "set": "Reusable cartridge is a great idea"
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
            "set": "Reusable cartridge is a great idea"
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
            "set": "good product  and device"
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
            "set": "good product  and device"
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
            "set": "it looks easy to use and easy to change cartridge - it has a large dose number indicator which is good"
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
            "set": "it looks easy to use and easy to change cartridge - it has a large dose number indicator which is good"
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
            "set": "appears to be a reusable device with cartridges.  good for the environment, but more difficult for patients with poor grip and dexterity"
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
            "set": "appears to be a reusable device with cartridges.  good for the environment, but more difficult for patients with poor grip and dexterity"
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
            "set": "these is a good product"
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
            "set": "these is a good product"
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
            "set": "Seems clear cut with good instructions"
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
            "set": "Seems clear cut with good instructions"
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
            "set": "Looks robust."
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
            "set": "Looks robust."
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
            "set": "looks good but would have to be cheaper than Product B"
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
            "set": "looks good but would have to be cheaper than Product B"
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
            "set": "Like the larger device - elderly people Struggle, also love the green to red colour change. Seems quite easy to use"
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
            "set": "Like the larger device - elderly people Struggle, also love the green to red colour change. Seems quite easy to use"
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
            "set": "patient freindly, easy to use"
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
            "set": "patient freindly, easy to use"
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
            "set": "Product B is difficult to use these improvments witll not help"
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
            "set": "Product B is difficult to use these improvments witll not help"
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
            "set": "It looks easy to use especially for older patients older patients struggle with small devices and with reading information regarding how many doses are left i feel this will be good. it also appears easy to use"
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
            "set": "It looks easy to use especially for older patients older patients struggle with small devices and with reading information regarding how many doses are left i feel this will be good. it also appears easy to use"
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
            "set": "Looks like some features are an advance with the easy looking cartridges the number counter and a bigger device - easier handling"
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
            "set": "Looks like some features are an advance with the easy looking cartridges the number counter and a bigger device - easier handling"
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
            "set": "Larger device Dose counter"
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
            "set": "Larger device Dose counter"
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
            "set": "Looks good would like to know about efficacy and cost"
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
            "set": "Looks good would like to know about efficacy and cost"
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
            "set": "more suitable for elderly population because of its size, indicator and probably easier to use in patients with limited dexterity"
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
            "set": "more suitable for elderly population because of its size, indicator and probably easier to use in patients with limited dexterity"
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
            "set": "need dexterity, lots of Condition Y patients are older and having other medical problems"
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
            "set": "need dexterity, lots of Condition Y patients are older and having other medical problems"
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
            "set": "Innovative. Might save cost  long term"
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
            "set": "Innovative. Might save cost  long term"
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
            "set": "Looks like a good device - easy to use and easy to see the acutations"
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
            "set": "Looks like a good device - easy to use and easy to see the acutations"
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
            "set": "It is simple and easy to use"
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
            "set": "It is simple and easy to use"
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
            "set": "simple and handy device"
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
            "set": "simple and handy device"
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
            "set": "Easy to use if no hand problems Simple to see doses"
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
            "set": "Easy to use if no hand problems Simple to see doses"
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
            "set": "easy to use for older patients higher capicity"
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
            "set": "easy to use for older patients higher capicity"
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
            "set": "A bit confusing for the elderly"
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
            "set": "A bit confusing for the elderly"
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
            "set": "Can envisage handling problems, particularly with elderly, arthritic or less dexterous patients"
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
            "set": "Can envisage handling problems, particularly with elderly, arthritic or less dexterous patients"
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
            "set": "It looks good and due to its large size may be better for less dextrous patients. I reserve judgement on the size of the numbers on the dose counter which i hope are more friendly than most."
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
            "set": "It looks good and due to its large size may be better for less dextrous patients. I reserve judgement on the size of the numbers on the dose counter which i hope are more friendly than most."
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
            "set": "i have seen this device or a device similar to this previously and had some concerns that although the deposition of the drug is very good the putting together of the device could be tricky for some or overwhelming."
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
            "set": "i have seen this device or a device similar to this previously and had some concerns that although the deposition of the drug is very good the putting together of the device could be tricky for some or overwhelming."
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
            "set": "Very similar to Product B but more environmentally friendly with good dose counter"
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
            "set": "Very similar to Product B but more environmentally friendly with good dose counter"
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
            "set": "Amazing this device is a game changer"
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
            "set": "Amazing this device is a game changer"
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
            "set": "Looks to be an improvement on current devices - larger, apparently, so easier to use. No doubt when device is empty so patient cannot continue using empty device. Partly reusable."
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
            "set": "Looks to be an improvement on current devices - larger, apparently, so easier to use. No doubt when device is empty so patient cannot continue using empty device. Partly reusable."
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
            "set": "complicated"
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
            "set": "complicated"
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
            "set": "Useful"
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
            "set": "Useful"
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
            "set": "Appears that it can be quite useful"
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
            "set": "Appears that it can be quite useful"
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
            "set": "I like the counter, looks easy to handle"
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
            "set": "I like the counter, looks easy to handle"
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
            "set": "Looks easy enough to use"
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
            "set": "Looks easy enough to use"
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
            "set": "Obviously not being able to handle it. Then it is difficult to judge size. However looks smaller so easier to carry. Is it easier to load a new cartridge, Product B is difficult so the pharmacists often do this. larger numbers is better for older people."
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
            "set": "Obviously not being able to handle it. Then it is difficult to judge size. However looks smaller so easier to carry. Is it easier to load a new cartridge, Product B is difficult so the pharmacists often do this. larger numbers is better for older people."
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
            "set": "Cost effective Refill easy to replace Dose counter"
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
            "set": "Cost effective Refill easy to replace Dose counter"
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
            "set": "still not straight forward to load and use"
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
            "set": "still not straight forward to load and use"
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
            "set": "easy to handle and manage, less dexterity needed."
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
            "set": "easy to handle and manage, less dexterity needed."
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
            "set": "Seems to be user friendly for patients"
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
            "set": "Seems to be user friendly for patients"
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
            "set": "sIMPLE AND EFFECTIVE - MAYBE TRICK FOR PATIENT WITH DEXTERITY ISSUES who are older"
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
            "set": "sIMPLE AND EFFECTIVE - MAYBE TRICK FOR PATIENT WITH DEXTERITY ISSUES who are older"
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
            "set": "This device appears to be much easier to use, with less packaging, less waste, and greater overall effectiveness of drug with simpler inhalation mechanism."
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
            "set": "This device appears to be much easier to use, with less packaging, less waste, and greater overall effectiveness of drug with simpler inhalation mechanism."
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
            "set": "i like this type of device, it looks very similar to the Product B and i know my patients like this device, i would be happy to recommend this system"
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
            "set": "i like this type of device, it looks very similar to the Product B and i know my patients like this device, i would be happy to recommend this system"
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
            "set": "A good option for patients when considering device device"
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
            "set": "A good option for patients when considering device device"
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
            "set": "Looks interesting  Modern Clear numbers Green"
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
            "set": "Looks interesting  Modern Clear numbers Green"
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
            "set": "too fiddly for the elderly"
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
            "set": "too fiddly for the elderly"
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
            "set": "Similar to current devices"
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
            "set": "Similar to current devices"
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
            "set": "Seems reasonable although no huge advantage over the alternatives"
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
            "set": "Seems reasonable although no huge advantage over the alternatives"
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
            "set": "Sounds good  Lots of advantages"
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
            "set": "Sounds good  Lots of advantages"
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
            "set": "Larger dose counter is very appealing as is the easy loading and automatic rejection of empty cartilage. Good for partially sighted patients as well as the elderly"
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
            "set": "Larger dose counter is very appealing as is the easy loading and automatic rejection of empty cartilage. Good for partially sighted patients as well as the elderly"
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
            "set": "it may be awkward to manipulate if you are arthritic. I like the idea of reusable though"
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
            "set": "it may be awkward to manipulate if you are arthritic. I like the idea of reusable though"
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
            "set": "Much improved. Easier to read dose counter. Like the fact its reuseable. Better for environment."
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
            "set": "Much improved. Easier to read dose counter. Like the fact its reuseable. Better for environment."
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
            "set": "Sensible idea re reducing waste but i worry Condition Y patients who are often older with co morbidities might struggle to change the cannister"
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
            "set": "Sensible idea re reducing waste but i worry Condition Y patients who are often older with co morbidities might struggle to change the cannister"
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
            "set": "It seems an improvement on the Product B. The initial assembly still seems awkward. I personally dislike the \"soft mist\" produced by the Product B, I would hope this would be an improvement on that."
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
            "set": "It seems an improvement on the Product B. The initial assembly still seems awkward. I personally dislike the \"soft mist\" produced by the Product B, I would hope this would be an improvement on that."
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
            "set": "This device appears to be a great improvement on the Product B Device. It seems as if it will be much easier to use, and  no chance of continuing to use it when empty."
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
            "set": "This device appears to be a great improvement on the Product B Device. It seems as if it will be much easier to use, and  no chance of continuing to use it when empty."
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
            "set": "usuable"
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
            "set": "usuable"
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
            "set": "good improvement to aid user and improve compliance"
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
            "set": "good improvement to aid user and improve compliance"
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
            "set": "Looks easy to use. Practices are driven by price. I am eco friendly so I like reusable. Condition Xtics do not like change when well controlled."
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
            "set": "Looks easy to use. Practices are driven by price. I am eco friendly so I like reusable. Condition Xtics do not like change when well controlled."
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
            "set": "appears simpler to use"
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
            "set": "appears simpler to use"
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
            "set": "dont think it will be easy for a lot of the elderly, less dextorous patients to use, looks difficult to change cylinder"
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
            "set": "dont think it will be easy for a lot of the elderly, less dextorous patients to use, looks difficult to change cylinder"
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
            "set": "Like that is is less waste, environmentally friendly,  Bigger device might be easier to use for elderly but younger people might prefer smaller device"
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
            "set": "Like that is is less waste, environmentally friendly,  Bigger device might be easier to use for elderly but younger people might prefer smaller device"
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
        ],
        [],
        [],
        [
          {
            "set": "Q8",
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
          {
            "set": "What are the primary strength(s) of the device you just reviewed?"
          }
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
            "set": "N",
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
          {
            "set": "N",
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
        ],
        [
          "Extended release",
          {
            "set": 0.79,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Extended release",
          {
            "set": 79,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Easy to use daily",
          {
            "set": 0.68,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Easy to use daily",
          {
            "set": 68,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Better for environment",
          {
            "set": 0.59,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Better for environment",
          {
            "set": 59,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Dose counter ",
          {
            "set": 0.79,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Dose counter ",
          {
            "set": 79,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Modern design",
          {
            "set": 0.18,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Modern design",
          {
            "set": 18,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Shape easy to handle",
          {
            "set": 0.49,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Shape easy to handle",
          {
            "set": 49,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Designed for multiple uses",
          {
            "set": 0.53,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Designed for multiple uses",
          {
            "set": 53,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Simple packet replacement",
          {
            "set": 0.7,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Simple packet replacement",
          {
            "set": 70,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Detaches when empty",
          {
            "set": 0.51,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Detaches when empty",
          {
            "set": 51,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Easy to switch",
          {
            "set": 0.42,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Easy to switch",
          {
            "set": 42,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Tough/ durable design",
          {
            "set": 0.27,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Tough/ durable design",
          {
            "set": 27,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Other (specify)",
          null,
          null,
          "Other (specify)",
          null
        ],
        [
          "None",
          {
            "set": 0.03,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "None",
          {
            "set": 3,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [],
        [],
        [
          {
            "set": "Q9v1",
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
          {
            "set": "What are the primary limitation(s) of the device you just reviewed?"
          }
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
            "set": "N",
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
          {
            "set": "N",
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
        ],
        [
          "Complicated to load packet and prime",
          {
            "set": 0.29,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Complicated to load packet and prime",
          {
            "set": 29,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Complicated to use every day",
          {
            "set": 0.06,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Complicated to use every day",
          {
            "set": 6,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Complicated to teach",
          {
            "set": 0.2,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Complicated to teach",
          {
            "set": 20,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Will have to re-train Product B patients ",
          {
            "set": 0.23,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Will have to re-train Product B patients ",
          {
            "set": 23,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Will make issuing prescriptions more complex (e.g. new device every 6 months)",
          {
            "set": 0.41,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Will make issuing prescriptions more complex (e.g. new device every 6 months)",
          {
            "set": 41,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Patients won't like change",
          {
            "set": 0.37,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Patients won't like change",
          {
            "set": 37,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Maintenance is difficult ",
          {
            "set": 0.39,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Maintenance is difficult ",
          {
            "set": 39,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Easy to lose / hard to replace ",
          {
            "set": 0.32,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Easy to lose / hard to replace ",
          {
            "set": 32,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Need to train GPs and Nurses",
          {
            "set": 0.37,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Need to train GPs and Nurses",
          {
            "set": 37,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          " (please specify)",
          null,
          null,
          " (please specify)",
          null
        ],
        [
          "I see no disadvantages",
          {
            "set": 0.16,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "I see no disadvantages",
          {
            "set": 16,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [],
        [],
        [],
        [
          {
            "set": "Q10",
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
          "Overall, on a scale of 1 to 10, how likely would you be to prescribe this new extended release Product B to an existing Product B user?"
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
            "set": "1 Not at all likely"
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
            "set": "1 Not at all likely"
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
            "set": "2"
          },
          {
            "set": 0.03,
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
            "set": "2"
          },
          {
            "set": 3,
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
            "set": "3"
          },
          {
            "set": 0.03,
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
            "set": "3"
          },
          {
            "set": 3,
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
            "set": "4"
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
            "set": "4"
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
            "set": "5"
          },
          {
            "set": 0.11,
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
            "set": "5"
          },
          {
            "set": 11,
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
            "set": "6"
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
            "set": "6"
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
            "set": "7"
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
            "set": "7"
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
            "set": "8"
          },
          {
            "set": 0.21,
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
            "set": "8"
          },
          {
            "set": 21,
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
            "set": "9"
          },
          {
            "set": 0.13,
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
            "set": "9"
          },
          {
            "set": 13,
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
            "set": "10 Extremely Likely"
          },
          {
            "set": 0.09,
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
            "set": "10 Extremely Likely"
          },
          {
            "set": 9,
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
        ],
        [],
        [],
        [],
        [
          {
            "set": "Q11",
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
          "On the same scale of 1 to 10, how likely would you be to prescribe this new extended release Product B device to a new patient?"
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
            "set": "[NA]"
          },
          {
            "set": 1,
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
            "set": "[NA]"
          },
          {
            "set": 100,
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
            "set": "1 Not at all likely"
          },
          null,
          null,
          {
            "set": "1 Not at all likely"
          },
          null
        ],
        [
          {
            "set": "2"
          },
          null,
          null,
          {
            "set": "2"
          },
          null
        ],
        [
          {
            "set": "3"
          },
          null,
          null,
          {
            "set": "3"
          },
          null
        ],
        [
          {
            "set": "4"
          },
          null,
          null,
          {
            "set": "4"
          },
          null
        ],
        [
          {
            "set": "5"
          },
          null,
          null,
          {
            "set": "5"
          },
          null
        ],
        [
          {
            "set": "6"
          },
          null,
          null,
          {
            "set": "6"
          },
          null
        ],
        [
          {
            "set": "7"
          },
          null,
          null,
          {
            "set": "7"
          },
          null
        ],
        [
          {
            "set": "8"
          },
          null,
          null,
          {
            "set": "8"
          },
          null
        ],
        [
          {
            "set": "9"
          },
          null,
          null,
          {
            "set": "9"
          },
          null
        ],
        [
          {
            "set": "10 Extremely Likely"
          },
          null,
          null,
          {
            "set": "10 Extremely Likely"
          },
          null
        ],
        [
          {
            "set": "6, or 7"
          },
          null,
          null,
          {
            "set": "6, or 7"
          },
          null
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
        ],
        [],
        [],
        [],
        [
          {
            "set": "Q12",
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
          "And on the same scale of 1 to 10, how likely would you be to prescribe this new extended release Product B device to an Condition X or Condition Y patient, using a different device currently (e.g. Product A, Product C, Product D, Product J, Product H etc.), but looking to switch?"
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
            "set": "[NA]"
          },
          {
            "set": 1,
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
            "set": "[NA]"
          },
          {
            "set": 100,
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
            "set": "1 Not at all likely"
          },
          null,
          null,
          {
            "set": "1 Not at all likely"
          },
          null
        ],
        [
          {
            "set": "2"
          },
          null,
          null,
          {
            "set": "2"
          },
          null
        ],
        [
          {
            "set": "3"
          },
          null,
          null,
          {
            "set": "3"
          },
          null
        ],
        [
          {
            "set": "4"
          },
          null,
          null,
          {
            "set": "4"
          },
          null
        ],
        [
          {
            "set": "5"
          },
          null,
          null,
          {
            "set": "5"
          },
          null
        ],
        [
          {
            "set": "6"
          },
          null,
          null,
          {
            "set": "6"
          },
          null
        ],
        [
          {
            "set": "7"
          },
          null,
          null,
          {
            "set": "7"
          },
          null
        ],
        [
          {
            "set": "8"
          },
          null,
          null,
          {
            "set": "8"
          },
          null
        ],
        [
          {
            "set": "9"
          },
          null,
          null,
          {
            "set": "9"
          },
          null
        ],
        [
          {
            "set": "10 Extremely Likely"
          },
          null,
          null,
          {
            "set": "10 Extremely Likely"
          },
          null
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
        ],
        [],
        [],
        [
          {
            "set": "Q13",
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
          {
            "set": "Now, thinking about the current Product B device, please can you rate the new extended release device on the following features?  Using a <b>7-point scale</b>, where;\n<b>-3 is significantly worse</b> than the current device\n<b>0 is the same</b> as the current device\n<b>+3 is significantly better </b>than the current device\n"
          }
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
            "set": "N",
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
          {
            "set": "N",
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
        ],
        [
          "Ease of daily use",
          {
            "set": 0.25,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Ease of daily use",
          {
            "set": 25,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Ease of packet and priming ",
          {
            "set": 0.24,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Ease of packet and priming ",
          {
            "set": 24,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Size",
          {
            "set": 0.24,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Size",
          {
            "set": 24,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Wastage",
          {
            "set": 0.47,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Wastage",
          {
            "set": 47,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Dose counter",
          {
            "set": 0.54,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Dose counter",
          {
            "set": 54,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ]
      ],
      "sheetViews": {
        "showGridLines": "0"
      }
    }
  ]
}

