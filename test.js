var Spreadsheet = require('./index.js');

/*
var spreadsheet = new Spreadsheet({'filepath':''});

spreadsheet.on('error',function() {
 console.log(spreadsheet.error);
}).parse(function() {
 console.log(spreadsheet.data);
});


var s2 = new Spreadsheet({'filepath':'ExitData.xls',
                          'sheet':'Sheet2'});
s2.parse(function() {
 console.log(s2.data);
});

*/

var s3 = new Spreadsheet({'filepath':'ExitData.csv'});
s3.parse(function() {
 console.log(s3.data);
});


/*
var s4 = new Spreadsheet({'filepath':'ExitData.csv',
                          'delimiter':','});
s4.parse(function() {
 console.log(s4.data);
});



var s5 = new Spreadsheet({'filepath':'ExitData.ods'});
s5.parse(function() {
 console.log(s5.data);
});

*/