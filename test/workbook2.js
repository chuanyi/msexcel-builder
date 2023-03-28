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
    var OUTFILE = './test/out/workbook2.xlsx';
    var TESTFILE = './test/files/workbook2.xlsx'

    // excelbuilder.defaults( {
    //   font: {
    //     sz: '12',
    //     name: 'Verdana'
    //   }
    // })
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

//
// var pojo = {
//   "worksheets": [
//     {
//       "name": "sheet1",
//       "sheetViews": {
//         "showGridLines": "0"
//       },
//       "cells": [
//         ["A", "B", "C", "D", "E"],
//         [],
//         [1, 2, three, 4, 5],
//         [6, 7, 8, 9, 10],
//         [11, 12, 13, 14, 15],
//
//       ]
//     },
//     {
//       "name": "sheet2",
//       "sheetViews": {
//         "showGridLines": "0"
//       },
//       "cells": [
//         ["A", "B", "C", "D", "E"],
//         [1, 2, three, 4, 5],
//         [6, 7, 8, 9, 10],
//         [11, 12, 13, 14, 15],
//       ]
//     }
//   ]
// }

var pojo = {

  "worksheets": [
    {
      "name": "Intro",
      "sheetViews": {
        "showGridLines": "0"
      },
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
            },
            "width": 32
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
            },
            "width": 32
          },
          {
            "set": "(Project Name)"
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
            "set": "12-05-2022 10:42 am"
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
            "set": "http://localhost:3000/api/v3/datasets/generations/app.html#filter/%7B%22global%22%3A%7B%7D%7D/options/%7B%22showMissing%22%3Atrue%2C%22showPercent%22%3Atrue%2C%22showFormat%22%3Atrue%2C%22showWeighted%22%3Atrue%7D",
            "l": {
              "target": "http://localhost:3000/api/v3/datasets/generations/app.html#filter/%7B%22global%22%3A%7B%7D%7D/options/%7B%22showMissing%22%3Atrue%2C%22showPercent%22%3Atrue%2C%22showFormat%22%3Atrue%2C%22showWeighted%22%3Atrue%7D"
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
          2511
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
          2511
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
            "set": "Q1-Q6"
          },
          {
            "set": "Gender and Generations Survey by Pew Research Center. Conducted Nov-Dec 2012 | File Release Date: 2014 January 7 "
          }
        ],
        [
          "",
          {
            "set": "Q11-Q17"
          },
          {
            "set": ""
          }
        ],
        [
          "",
          {
            "set": "q15"
          },
          {
            "set": "Do you think young adults & older adults are similar or different today in terms of  ?"
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
      "name": "Q1Q6",
      "sheetViews": {
        "showGridLines": "0"
      },
      "cells": [
        [],
        [],
        [],
        [
          {
            "set": "Q1-Q6",
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
            "set": "Gender and Generations Survey by Pew Research Center. Conducted Nov-Dec 2012 | File Release Date: 2014 January 7 "
          }
        ],
        [],
        [],
        [],
        [],
        [
          {
            "set": "q1",
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
          "Q.1 Generally, how would you say things are these days in your life - would you say that you are very happy, pretty happy, or not too happy?"
        ],
        [
          {
            "set": "vs. q1",
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
            "set": "Q.1 Generally, how would you say things are these days in your life - would you say that you are very happy, pretty happy, or not too happy?"
          }
        ],
        [
          {
            "set": "q1",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Very happy",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Pretty happy",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Not too happy",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Don't know/Refused (VOL.)",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          null,
          {
            "set": "q1",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Very happy",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Pretty happy",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Not too happy",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Don't know/Refused (VOL.)",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          }
        ],
        [
          {
            "set": "Very happy"
          },
          {
            "set": 0.305057745917961,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 1,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          null,
          null,
          null,
          {
            "set": "Very happy"
          },
          {
            "set": 766,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 766,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          null,
          null
        ],
        [
          {
            "set": "Pretty happy"
          },
          {
            "set": 0.5085623257666269,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          {
            "set": 1,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          null,
          null,
          {
            "set": "Pretty happy"
          },
          {
            "set": 1277,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          {
            "set": 1277,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          null
        ],
        [
          {
            "set": "Not too happy"
          },
          {
            "set": 0.15412186379928317,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          null,
          {
            "set": 1,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          null,
          {
            "set": "Not too happy"
          },
          {
            "set": 387,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          null,
          {
            "set": 387,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null
        ],
        [
          {
            "set": "Don't know/Refused (VOL.)"
          },
          {
            "set": 0.03225806451612903,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          null,
          null,
          {
            "set": 1,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          {
            "set": "Don't know/Refused (VOL.)"
          },
          {
            "set": 81,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          null,
          null,
          {
            "set": 81,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
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
            },
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
            "set": 2511,
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
            "set": 766,
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
            "set": 1277,
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
            "set": 387,
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
            "set": 81,
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
            "set": 2511,
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
            "set": 766,
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
            "set": 1277,
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
            "set": 387,
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
            "set": 81,
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
            "set": "q2",
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
          "Q.2 How would you rate your own health in general these days? Would you say your health is excellent, good, only fair, or poor? "
        ],
        [
          {
            "set": "vs. q1",
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
            "set": "Q.1 Generally, how would you say things are these days in your life - would you say that you are very happy, pretty happy, or not too happy?"
          }
        ],
        [
          {
            "set": "q1",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Very happy",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Pretty happy",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Not too happy",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Don't know/Refused (VOL.)",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          null,
          {
            "set": "q1",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Very happy",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Pretty happy",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Not too happy",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Don't know/Refused (VOL.)",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          }
        ],
        [
          {
            "set": "Excellent"
          },
          {
            "set": 0.2536837913181999,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.412532637075718,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.21299921691464369,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.09819121447028424,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.13580246913580246,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          {
            "set": "Excellent"
          },
          {
            "set": 637,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 316,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 272,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 38,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 11,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "set": "Good"
          },
          {
            "set": 0.5173237753882916,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.46736292428198434,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.5880971025841817,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.3798449612403101,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.5308641975308642,
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
            "set": "Good"
          },
          {
            "set": 1299,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 358,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 751,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 147,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 43,
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
            "set": "Only fair"
          },
          {
            "set": 0.1776184786937475,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.10574412532637076,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.16914643696162882,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.34108527131782945,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.20987654320987653,
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
            "set": "Only fair"
          },
          {
            "set": 446,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 81,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 216,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 132,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 17,
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
            "set": "Poor"
          },
          {
            "set": 0.04858622062923138,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.01174934725848564,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.02975724353954581,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.1731266149870801,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.09876543209876543,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          {
            "set": "Poor"
          },
          {
            "set": 122,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 9,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 38,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 67,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 8,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
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
            },
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
            "set": "Don't know/Refused (VOL.)"
          },
          {
            "set": 0.0027877339705296694,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.0026109660574412533,
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
            "set": 0.007751937984496124,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.024691358024691357,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          {
            "set": "Don't know/Refused (VOL.)"
          },
          {
            "set": 7,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
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
          },
          null,
          {
            "set": 3,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 2,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
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
            },
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
            "set": 2511,
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
            "set": 766,
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
            "set": 1277,
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
            "set": 387,
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
            "set": 81,
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
            "set": 2511,
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
            "set": 766,
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
            "set": 1277,
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
            "set": 387,
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
            "set": 81,
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
            "set": "q3",
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
          "Q.3 How would you describe your household’s financial situation? Would you say you [READ; DO NOT RANDOMIZE]"
        ],
        [
          {
            "set": "vs. q1",
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
            "set": "Q.1 Generally, how would you say things are these days in your life - would you say that you are very happy, pretty happy, or not too happy?"
          }
        ],
        [
          {
            "set": "q1",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Very happy",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Pretty happy",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Not too happy",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Don't know/Refused (VOL.)",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          null,
          {
            "set": "q1",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Very happy",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Pretty happy",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Not too happy",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          },
          {
            "set": "Don't know/Refused (VOL.)",
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
            "wrap": true,
            "align": "wrap",
            "width": 12
          }
        ],
        [
          {
            "set": "Live comfortably"
          },
          {
            "set": 0.36837913181999205,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.54177545691906,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.35160532498042285,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.1111111111111111,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.2222222222222222,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          {
            "set": "Live comfortably"
          },
          {
            "set": 925,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 415,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 449,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 43,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 18,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "set": "Meet your basic expenses with a little left over for extras"
          },
          {
            "set": 0.2923138191955396,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.24673629242819844,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.34690681284259983,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.20671834625322996,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.2716049382716049,
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
            "set": "Meet your basic expenses with a little left over for extras"
          },
          {
            "set": 734,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 189,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 443,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 80,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 22,
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
            "set": "Just meet your basic expenses, or"
          },
          {
            "set": 0.2262046993229789,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.1514360313315927,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.2255285826155051,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.35917312661498707,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.30864197530864196,
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
            "set": "Just meet your basic expenses, or"
          },
          {
            "set": 568,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 116,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 288,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 139,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 25,
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
            "set": "Don’t even have enough to meet basic expenses"
          },
          {
            "set": 0.09199522102747909,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.03655352480417755,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.060297572435395456,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.2997416020671835,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.12345679012345678,
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
            "set": "Don’t even have enough to meet basic expenses"
          },
          {
            "set": 231,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 28,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 77,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "fill": {
              "type": "solid",
              "fgColor": "FFEEEEEE",
              "bgColor": "-"
            },
            "font": {
              "bold": "-",
              "color": "FF666666",
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
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 116,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 10,
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
            "set": "[VOL. DO NOT READ] Don’t know/Refused"
          },
          {
            "set": 0.021107128634010354,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.02349869451697128,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.015661707126076743,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.023255813953488372,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 0.07407407407407407,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@",
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
            },
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          null,
          {
            "set": "[VOL. DO NOT READ] Don’t know/Refused"
          },
          {
            "set": 53,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": {
                "style": "thin",
                "color": "888888"
              },
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 18,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
          },
          {
            "set": 20,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
            "border": {
              "left": "-",
              "right": "-",
              "top": "-",
              "bottom": "-"
            }
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
          },
          {
            "set": 6,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@",
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
            },
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
            "set": 2511,
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
            "set": 766,
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
            "set": 1277,
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
            "set": 387,
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
            "set": 81,
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
            "set": 2511,
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
            "set": 766,
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
            "set": 1277,
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
            "set": 387,
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
            "set": 81,
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
    },
    {
      "name": "q15",
      "sheetViews": {
        "showGridLines": "0"
      },
      "cells": [
        [],
        [],
        [
          {
            "set": "q15",
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
            "set": "Do you think young adults & older adults are similar or different today in terms of  ?"
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
            "set": 2511,
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
            "set": 2511,
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
          "Their moral values ",
          {
            "set": 0.2118677817602549,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Their moral values ",
          {
            "set": 532,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "The way they use the internet & other technology ",
          {
            "set": 0.12345679012345678,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "The way they use the internet & other technology ",
          {
            "set": 310,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Attitudes about racial & ethnic makeup of the country ",
          {
            "set": 0.21983273596176822,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Attitudes about racial & ethnic makeup of the country ",
          {
            "set": 552,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
          }
        ],
        [
          "Importance of family ",
          {
            "set": 0.35643170051772205,
            "numberFormat": "\"\"0.0%;\"\"\\-0.0%;\"\"\\—;@"
          },
          null,
          "Importance of family ",
          {
            "set": 895,
            "numberFormat": "\"\"0;\"\"\\-0;\"\"\\-\\-;@"
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
      "name": "Q11Q17",
      "sheetViews": {
        "showGridLines": "0"
      },
      "cells": [
        [],
        [],
        [],
        [
          {
            "set": "Q11-Q17",
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
          null
        ],
        [],
        [],
        [],
        [],
        [
          {
            "set": "q11",
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
          "Q.11 In terms of your values & beliefs, how much do you feel you have in common with other members of YOUR age group or generation? Would you say you have [READ] with other members of your generation?"
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
            "set": "A great deal in common"
          },
          {
            "set": 0.363600159299084,
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
            "set": "A great deal in common"
          },
          {
            "set": 913,
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
            "set": "Some things in common"
          },
          {
            "set": 0.5173237753882916,
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
            "set": "Some things in common"
          },
          {
            "set": 1299,
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
            "set": "Little in common"
          },
          {
            "set": 0.09279171644763043,
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
            "set": "Little in common"
          },
          {
            "set": 233,
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
            "set": "Nothing at all in common"
          },
          {
            "set": 0.011947431302270013,
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
            "set": "Nothing at all in common"
          },
          {
            "set": 30,
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
            "set": "Don???t know/Refused (VOL.)"
          },
          {
            "set": 0.014336917562724014,
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
            "set": "Don???t know/Refused (VOL.)"
          },
          {
            "set": 36,
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
            "set": 2511,
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
            "set": 2511,
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
            "set": "q12",
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
          "Q.12 (In general/And), how well do you think YOUNG adults understand the problems and concerns of OLDER adults?  Would you say they understand them very well, somewhat well, not too well or not well at all?"
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
            "set": "Very well"
          },
          {
            "set": 0.04221425726802071,
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
            "set": "Very well"
          },
          {
            "set": 106,
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
            "set": "Somewhat well"
          },
          {
            "set": 0.29669454400637196,
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
            "set": "Somewhat well"
          },
          {
            "set": 745,
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
            "set": "Not too well"
          },
          {
            "set": 0.4679410593389088,
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
            "set": "Not too well"
          },
          {
            "set": 1175,
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
            "set": "Not well at all"
          },
          {
            "set": 0.17801672640382318,
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
            "set": "Not well at all"
          },
          {
            "set": 447,
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
            "set": "Don???t know/Refused (VOL.)"
          },
          {
            "set": 0.015133412982875348,
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
            "set": "Don???t know/Refused (VOL.)"
          },
          {
            "set": 38,
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
            "set": 2511,
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
            "set": 2511,
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
            "set": "q13",
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
          "Q.13 (In general/And), how well do you think OLDER adults understand the problems and concerns of YOUNG adults?  Would you say they understand them very well, somewhat well, not too well or not well at all?"
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
            "set": "Very well"
          },
          {
            "set": 0.18876941457586618,
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
            "set": "Very well"
          },
          {
            "set": 474,
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
            "set": "Somewhat well"
          },
          {
            "set": 0.4994026284348865,
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
            "set": "Somewhat well"
          },
          {
            "set": 1254,
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
            "set": "Not too well"
          },
          {
            "set": 0.24611708482676226,
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
            "set": "Not too well"
          },
          {
            "set": 618,
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
            "set": "Not well at all"
          },
          {
            "set": 0.050577459179609714,
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
            "set": "Not well at all"
          },
          {
            "set": 127,
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
            "set": "Don???t know/Refused (VOL.)"
          },
          {
            "set": 0.015133412982875348,
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
            "set": "Don???t know/Refused (VOL.)"
          },
          {
            "set": 38,
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
            "set": 2511,
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
            "set": 2511,
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
            "set": "q14",
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
          "Q.14 How likely is it that, at some point in your life, you will be responsible for caring for an aging parent or another elderly family member?  Do you think it is very likely, somewhat likely, not too likely, or not at all likely?"
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
            "set": "Very likely"
          },
          {
            "set": 0.40700915969733176,
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
            "set": "Very likely"
          },
          {
            "set": 1022,
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
            "set": "Somewhat likely"
          },
          {
            "set": 0.19195539625647154,
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
            "set": "Somewhat likely"
          },
          {
            "set": 482,
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
            "set": "Not too likely"
          },
          {
            "set": 0.1015531660692951,
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
            "set": "Not too likely"
          },
          {
            "set": 255,
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
            "set": "Not at all likely"
          },
          {
            "set": 0.16009557945041816,
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
            "set": "Not at all likely"
          },
          {
            "set": 402,
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
            "set": "Depends (VOL.)"
          },
          {
            "set": 0.0011947431302270011,
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
            "set": "Depends (VOL.)"
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
            "set": "Have already done this/Am currently doing this (VOL.)"
          },
          {
            "set": 0.1306252489048188,
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
            "set": "Have already done this/Am currently doing this (VOL.)"
          },
          {
            "set": 328,
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
            "set": "Don't know/Refused (VOL.)"
          },
          {
            "set": 0.007566706491437674,
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
            "set": "Don't know/Refused (VOL.)"
          },
          {
            "set": 19,
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
            "set": 2511,
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
            "set": 2511,
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
            "set": "q16",
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
          null
        ],
        [],
        [],
        [],
        [],
        [
          {
            "set": "q16a",
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
          "Q.16a You said young adults and older adults were different in terms of their moral values. In your opinion, who has the better moral values:  Young adults or older adults?"
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
            "set": 0.24452409398645958,
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
            "set": 614,
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
            "set": "Young adults"
          },
          {
            "set": 0.03544404619673437,
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
            "set": "Young adults"
          },
          {
            "set": 89,
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
            "set": "Older adults"
          },
          {
            "set": 0.6451612903225806,
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
            "set": "Older adults"
          },
          {
            "set": 1620,
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
            "set": "Neither better nor worse just different/mixed (VOL.)"
          },
          {
            "set": 0.04340900039824771,
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
            "set": "Neither better nor worse just different/mixed (VOL.)"
          },
          {
            "set": 109,
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
            "set": "Don't know/Refused (VOL.)"
          },
          {
            "set": 0.0314615690959777,
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
            "set": "Don't know/Refused (VOL.)"
          },
          {
            "set": 79,
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
            "set": 2511,
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
            "set": 2511,
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
            "set": "q16b",
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
          "Q.16b You also said young adults and older adults were different in terms of their attitudes about the changing racial and ethnic makeup of the country. In your opinion, who has the better attitudes: Young adults or older adults?"
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
            "set": 0.2843488649940263,
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
            "set": 714,
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
            "set": "Young adults"
          },
          {
            "set": 0.47192353643966545,
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
            "set": "Young adults"
          },
          {
            "set": 1185,
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
            "set": "Older adults"
          },
          {
            "set": 0.2019115890083632,
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
            "set": "Older adults"
          },
          {
            "set": 507,
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
            "set": "Neither better nor worse just different/mixed (VOL.)"
          },
          {
            "set": 0.019912385503783353,
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
            "set": "Neither better nor worse just different/mixed (VOL.)"
          },
          {
            "set": 50,
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
            "set": "Don't know/Refused (VOL.)"
          },
          {
            "set": 0.02190362405416169,
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
            "set": "Don't know/Refused (VOL.)"
          },
          {
            "set": 55,
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
            "set": 2511,
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
            "set": 2511,
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
            "set": "q17",
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
            "set": "Do you think this is a responsibility or is it not really a responsibility? - "
          }
        ],
        [],
        [],
        [],
        [],
        [
          {
            "set": "q17a",
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
          "Q.17a Parents providing financial assistance to an adult child if he or she needs it"
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
            "set": "Responsibility"
          },
          {
            "set": 0.5264834727200318,
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
            "set": "Responsibility"
          },
          {
            "set": 1322,
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
            "set": "Not really a responsibility"
          },
          {
            "set": 0.4336917562724014,
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
            "set": "Not really a responsibility"
          },
          {
            "set": 1089,
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
            "set": "Don't know/Refused (VOL.)"
          },
          {
            "set": 0.039824771007566706,
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
            "set": "Don't know/Refused (VOL.)"
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
          "N",
          {
            "set": 2511,
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
            "set": 2511,
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
            "set": "q17b",
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
          "Q.17b Adult children providing financial assistance to an elderly parent if he or she needs it"
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
            "set": "Responsibility"
          },
          {
            "set": 0.7307845479888491,
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
            "set": "Responsibility"
          },
          {
            "set": 1835,
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
            "set": "Not really a responsibility"
          },
          {
            "set": 0.24731182795698925,
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
            "set": "Not really a responsibility"
          },
          {
            "set": 621,
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
            "set": "Don't know/Refused (VOL.)"
          },
          {
            "set": 0.02190362405416169,
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
            "set": "Don't know/Refused (VOL.)"
          },
          {
            "set": 55,
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
            "set": 2511,
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
            "set": 2511,
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

