var fs = require('fs');
var xlsx = require('xlsx');
var xls = require('xlsjs');
var streamifier = require('streamifier');
var EventEmitter = require('events').EventEmitter;

var MyClass = function(config){
 config = config || {}; // Holds any config data if its passed in
 this.name = config.name || 'Untitled';
 this.sheet = config.sheet || '';
 this.error = null;
 this.data = new Array();
 this.delimiter = config.delimiter || ',';

 var me = this;
 var callbackfn;
 var buf;
 var stream;
 var outputObj;
 var filepath = config.filepath || '';

 this.parse = function(callback) {
  callbackfn=callback;
  if (filepath=='') {
   this.error='No File Specified';
   this.emit('error');
   return;
  }
  try {
   buf = fs.readFileSync(filepath);
  } catch (e) {
   this.error='Invalid Filename';
   this.emit('error');
  }
  switch (buf.toString("base64",0,4)) {
   case 'UEsDBA==':
    this.data = parseExcelObj(xlsx.read(buf),this.sheet);
    callbackfn();
    break;
   case '0M8R4A==':
    this.data = parseExcelObj(xls.read(buf),this.sheet);
    callbackfn();
    break;
   default:
    this.data = CSVToArray(buf.toString('ascii'));
    callbackfn();
  }
 };
};



function parseExcelObj(obj,sheetname) {
 var outputArr=new Array();
 if (sheetname=='') {
  sheetname=obj.SheetNames[0];
 }
 var range;
 try {
  range=obj.Sheets[sheetname]['!ref'].split(':');
 } catch (e) {
  return outputArr;
 }

 for (var i in range) {
  range[i]=excelCoordToMatrix(range[i]);
 }
 for (var i=0;i <= range[1].ycoord;i++) {
  outputArr.push(new Array());
 }
 Object.keys(obj.Sheets[sheetname]).forEach(function(key) {
  var coords;
  if (key != '!ref' && key != '!range') {
   coords = excelCoordToMatrix(key);
   outputArr[coords.ycoord][coords.xcoord]=obj.Sheets[sheetname][key].w;
  }
 });
 return outputArr;
}


function excelCoordToMatrix(xlsCoord) {  // ex. converts D5 to [3,4]
 var xcoord = -1;
 var ycoord = -1;
 for (var i=0;i < xlsCoord.length; i++) {
  if (xlsCoord.charCodeAt(i) >= 65) {
   xcoord = 26*(xcoord+1) + xlsCoord.charCodeAt(i)-65;
  } else {
   ycoord = 10*(ycoord+1) + xlsCoord.charCodeAt(i)-49;
  }
 }
 return {'xcoord':xcoord,'ycoord':ycoord};
}



function CSVToArray( strData, strDelimiter ){
	// Check to see if the delimiter is defined. If not,
	// then default to comma.
	strDelimiter = (strDelimiter || ",");

	// Create a regular expression to parse the CSV values.
	var objPattern = new RegExp(
		(
			// Delimiters.
			"(\\" + strDelimiter + "|\\r?\\n|\\r|^)" +

			// Quoted fields.
			"(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|" +

			// Standard fields.
			"([^\"\\" + strDelimiter + "\\r\\n]*))"
		),
		"gi"
		);


	// Create an array to hold our data. Give the array
	// a default empty first row.
	var arrData = [[]];

	// Create an array to hold our individual pattern
	// matching groups.
	var arrMatches = null;


	// Keep looping over the regular expression matches
	// until we can no longer find a match.
	while (arrMatches = objPattern.exec( strData )){

		// Get the delimiter that was found.
		var strMatchedDelimiter = arrMatches[ 1 ];

		// Check to see if the given delimiter has a length
		// (is not the start of string) and if it matches
		// field delimiter. If id does not, then we know
		// that this delimiter is a row delimiter.
		if (
			strMatchedDelimiter.length &&
			(strMatchedDelimiter != strDelimiter)
			){

			// Since we have reached a new row of data,
			// add an empty row to our data array.
			arrData.push( [] );

		}


		// Now that we have our delimiter out of the way,
		// let's check to see which kind of value we
		// captured (quoted or unquoted).
		if (arrMatches[ 2 ]){

			// We found a quoted value. When we capture
			// this value, unescape any double quotes.
			var strMatchedValue = arrMatches[ 2 ].replace(
				new RegExp( "\"\"", "g" ),
				"\""
				);

		} else {

			// We found a non-quoted value.
			var strMatchedValue = arrMatches[ 3 ];

		}


		// Now that we have our value string, let's add
		// it to the data array.
		arrData[ arrData.length - 1 ].push( strMatchedValue );
	}

	// sometimes csv files record an extra blank line at the end
	if (arrData.length > 0) {
		if (arrData[arrData.length-1].length == 1) {
			if (arrData[arrData.length-1][0] == '') {
				arrData.pop();
			}
		}
	}

	// Return the parsed data.
	return( arrData );
}


require('util').inherits(MyClass, EventEmitter);
module.exports = MyClass;