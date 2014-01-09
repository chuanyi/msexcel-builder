var excelbuilder = require('../lib/msexcel-builder.js')

var workbook = excelbuilder.createWorkbook('./', 'sample.xlsx')

// Create a new worksheet with 10 columns and 12 rows
var sheet1 = workbook.createSheet('Sheet1', 10, 12);

// test write number
sheet1.set(1, 1, "String");
sheet1.set(1, 1, "Demo");
sheet1.set(2, 2, 5);
sheet1.set(2, 3, 5.0);

// Save it
workbook.save(function(err){
	if (err) 
	  workbook.cancel();
	else
	{
	  console.log('congratulations, your workbook created');
	}
});