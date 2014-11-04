parse-spreadsheet
=================

For use in node.js.

Easily parse a spreadsheet of .xls, .xlsx, .ods, .csv or .txt format either synchronously or asynchronously and return a matrix of string values corresponding to the table cell values.

If inputting an .xls, .xlsx, or .ods file, an optional value titled 'sheet' may be included.  If not, then the function returns the data of the first sheet in the workbook.

If inputting an ascii file, an optional value titled 'delimiter' may be included.  If not, the value defaults to a comma.

If inputting a buffer instead of a filename, you only have the synchronous option available to receive the return data.

Usage Examples:

```
var fs = require('fs');
var Spreadsheet = require('./index.js');
var data;

data = Spreadsheet.parse({filepath:'./file.xlsx'});

Spreadsheet.parse({filepath:'./file.csv'},function(err,retdata) {
  data = retdata;
});

var buffer = fs.readFileSync('./file.ods');
data = Spreadsheet.parse({filebuffer:buffer,sheet:'Sheet1'});

Spreadsheet.parse({filepath:'./file.txt',delimiter:'\t'},function(err,retdata) {
  data = retdata;
});
```