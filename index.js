#! /usr/bin/env node
'use strict';

// The goal of this module is to split an excel file into multiple files based
// on a number of rows in a worksheet
const xlsx = require('async-xlsx');
const xlsx2 = require('xlsx');
const program = require('commander');
const _ = require('lodash');

// Steps
// 1. Receive args (file loc, worksheet to use, line count, output loc)
// -- Note that file paths must use process.cwd as this is an npm module

program
  .version('0.0.1')
  .option('-J, --settingsJSON <settingsJSON>', 'settings JSON file')
  .option('-S, --source <source>', 'data source')
  .option('-N, --lines <lines>', 'line count')
  .option('-O, --output <output>', 'output folder')
  .option('-W, --worksheet <worksheet>', 'worksheet')
  .parse(process.argv);

const splitSettings = require(process.cwd() + '/' + program.settingsJSON);

function Workbook() {
	if(!(this instanceof Workbook)) return new Workbook();
	 this.SheetNames = [];
	 this.Sheets = {};
}

  function datenum(v, date1904) {
    if (date1904) v += 1462;
    var epoch = Date.parse(v);
    return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
  }

  function numdate(v) {
  	var startDate = xlsx2.SSF.parse_date_code(v);
  	var val = new Date();
  	if(startDate !== null) {
    	val = startDate.m + '/' + startDate.d + '/' + startDate.y.toString().slice(2, 4);
    }
  	return val;
  }

  function sheet_from_array_of_arrays(data, opts) {
  var ws = {};
  var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
  for(var R = 0; R != data.length; ++R) {
  	for(var C = 0; C != data[R].length; ++C) {
  		if(range.s.r > R) range.s.r = R;
  		if(range.s.c > C) range.s.c = C;
  		if(range.e.r < R) range.e.r = R;
  		if(range.e.c < C) range.e.c = C;
  		var cell = {v: data[R][C] };
  		if(cell.v == null) continue;
  		var cell_ref = xlsx2.utils.encode_cell({c:C,r:R});

  		if(typeof cell.v === 'number') cell.t = 'n';
  		else if(typeof cell.v === 'boolean') cell.t = 'b';
  		else if(cell.v instanceof Date) {
  			cell.t = 'd';
        cell.z = 'dd/mm/yy';
  			cell.v = numdate(cell.v);
  		}
  		else cell.t = 's';

  		ws[cell_ref] = cell;
  	}
  }
  if (range.s.c < 10000000) ws['!ref'] = xlsx2.utils.encode_range(range);
    return ws;
  }

const tweakData = (sheetName, dataArray) => {
  // Expects a 2d array
  const newData = JSON.parse(JSON.stringify(dataArray));
  for (let zz = 0; zz < splitSettings.forceTypes.length; ++zz) {
    const forceType = splitSettings.forceTypes[zz];
    if (forceType.sheetName === sheetName) {
      // Only apply to matching sheets, ya goober
      for (let yy = forceType.startIndex || 0; yy < newData.length; ++yy) {
        const row = newData[yy];
        console.log(row);
        row[forceType.index] = numdate(row[forceType.index]);
      }
    }
  }
  return newData;
};
// 2. Load file into memory and create copy of workbook
return xlsx.parseFileAsync((process.cwd() + '/' + program.source), {}, workbookData => {
  const wbData = JSON.parse(JSON.stringify(workbookData));
  // Result is a 2D array with an object with name/data props for each worksheet
   //console.log(JSON.stringify(wbData, 0, 2));

  // 3. Isolate data array of rows on target spreadsheet
  const dataSet = _.find(wbData, { name: program.worksheet }).data;
  const sheetNames = _.map(wbData, 'name');
  const loadedSheets = {};
  _.map(sheetNames, sheetname => {
    const sheet = _.find(wbData, { name: sheetname });
    // Apply force types
    sheet.data = tweakData(sheetname, sheet.data);
    console.log(JSON.stringify(sheet, 0, 2));
    loadedSheets[sheetname] = sheet_from_array_of_arrays(sheet.data);
  });
  const lineCount = Number(program.lines);
  // 4a. Loop through the array of rows to build subsets
  for (let startIndex = 0; startIndex < (dataSet.length % lineCount === 0 ? dataSet.length / lineCount : dataSet.length / lineCount); ++startIndex) {
    let subset = dataSet.slice((startIndex * lineCount) + 1, (startIndex + 1) * lineCount);
    // Add headers into data
    subset = tweakData(program.worksheet, subset);
    subset.unshift(dataSet[0]);
    const workBook = { SheetNames: sheetNames, Sheets: loadedSheets };
    // 4b. Add each subset to a temporary copy of the workbook
    workBook.Sheets[program.worksheet] = sheet_from_array_of_arrays(subset);
    // 4c. Save the temporary workbook
    xlsx2.writeFile(workBook, process.cwd() + '/' + program.output + '/' + startIndex.toString() + '.xlsx');
  }
});
