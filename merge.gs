"use strict";
/*                                     MERGE GOOGLE SHEETS
*
*                                         Coert Vonk
*                                        February 2019
*
*                         https://github.com/cvonk/gas-sheets-merge
* USE:
*
* The names of the src spreadsheet, src sheet and dst sheet and at the end of this code.
* The onMerge() function is involved through the menu bar in Google Sheets (CUSTOM > Merge).
* Even if you don't ready anything else, please read through the example below.
*
* DEFINITIONS:
*
* The first row of the src/dst sheet should contain labels for each column.
* The left-most label in the dst sheet is considered the keyLabel, and the values in that
* column are considered "keys". These keys are used to match rows between the src and dst
* sheets.
* 
* FUNCTIONALITY:
*
* The function doMerge() should be called from the destination sheet.
* It walks the src sheet row-by-row.  If the key doesn't already exist in the dst, a new
* row is added to the dst.  In the dst sheet, the row with the corresponding key is populated
* with values from the src sheet.
* 
* EXAMPLE:
*
* src sheet                                    dst sheet
* +--------+------------+----------------+     +------------+--------+--------+
* | Class  | Animal     | Family         |     | Animal     | Class  | Food   |
* +--------+------------+----------------+     +------------+--------+--------+
* | mammal | moose      | cervidae       |     | moose      | mammal | plants |
* | mammal | dolphin    | delphinidae    |     | deer       | mammal | plants |
* | fish   | pufferfish | tetraodontidae |     +------------+--------+--------+
* +--------+------------+----------------+     
*
* The keyLabel is "Animal" since the top-left entry in the dst sheet.  In this example, the
* function doMerge() will:
*
*  1. empty the columnns in the dst sheet      +------------+--------+--------+
*     that will be sourced from the src sheet, | Animal     | Class  | Food   |
*     except for the first column that         +------------+--------+--------+
*     contains the key values;                 | moose      |        | plants |
*                                              | deer       |        | plants |
*                                              +------------+--------+--------+
*
*  2. appends rows for which there is no       +------------+--------+--------+
*     corresponding key value in the dst.      | Animal     | Class  | Food   |
*                                              +------------+--------+--------+
*                                              | moose      |        | plants |
*                                              | deer       |        | plants |
*                                              | dophin     |        |        |
*                                              | pufferfish |        |        |
*                                              +------------+--------+--------+
*
*  3. copy the src columns that also exist     +------------+--------+--------+
*     in the dst sheet.                        | Animal     | Class  | Food   |
*                                              +------------+--------+--------+
*                                              | moose      | mammal | plants |
*                                              | deer       |        | plants |
*                                              | dophin     | mammal |        |
*                                              | pufferfish | fish   |        |
*                                              +------------+--------+--------+
* 
*  4. Note: the empty class value for animal "deer" indicates that it (no longer) exists
*     in the src.  One could delete or hide the corresponding row.  Also, the Food for
*     "dophin" and "pufferfish" is still blank.
*
* LEGAL:
*
* (c) Copyright 2019 by Coert Vonk
*
* This program is free software: you can redistribute it and/or modify
* it under the terms of the GNU General Public License as published by
* the Free Software Foundation, either version 3 of the License, or
* (at your option) any later version.
*
* This program is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
* GNU General Public License for more details.
*
* You should have received a copy of the GNU General Public License
* along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

// add CUSTOM to menu bar

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('CUSTOM')
  .addItem('Merge', 'onMerge')
  .addToUi();
}

// update dst sheet based on src sheet

var OnMerge = {};
(function() {
  
  // open named spreadsheets from Google Drive
  
  this.openSpreadsheetByName = function(fname) {
    
    var files = DriveApp.searchFiles('mimeType = "' + MimeType.GOOGLE_SHEETS + '" and title contains "' + fname + '"');
    while (files.hasNext()) {
      return SpreadsheetApp.open(files.next());
    }
    return null;
  }
  
  this.openNamedSpreadsheet = function(spreadsheetName) {
    
    var spreadsheet = this.openSpreadsheetByName(spreadsheetName);
    if (spreadsheet == undefined) {
      throw "Spreadsheet \"" + spreadsheetName + "\" not found on gDrive";
    }
    return spreadsheet;  
  }
  
  this.openSheet = function(sheetName, minNrOfDataRows, requireLabelInA1, spreadsheetName) {
    
    var spreadsheet;
    if (spreadsheetName == undefined) {
      spreadsheetName = "active spreadsheet";
      spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    } else {
      spreadsheet = this.openNamedSpreadsheet(spreadsheetName);
    }
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet == undefined) {
      throw "Sheet \"" + sheetName + "\" not found in spreadsheet \"" + spreadsheetName + "\"";
    }
    if (requireLabelInA1 && sheet.getRange("A1").isBlank()) {
      throw "Sheet \"" + sheetName + "\" (part of " + spreadsheetName + ") doesn't have a keyLabel in A1";
    }
    var dbg = sheet.getDataRange().getNumRows();
    if (sheet.getDataRange().getNumRows() < minNrOfDataRows) {
      throw "Sheet \"" + sheetName + "\" (part of " + spreadsheetName + ") doesn't at least " + minNrOfDataRows + " rows";
    }
    return sheet;
  }
  
  this.srcSheetGetValues = function(sheet) {
    
    var range = sheet.getDataRange();
    return range.getValues();
  }
  
  // returns 1-dim array with header values
  
  this.dstSheetGetHeaderValues = function(dstSheet) {
    
    var dstHeaderNumOfCols = dstSheet.getDataRange().getNumColumns();     
    return dstSheet.getRange(1, 1, 1, dstHeaderNumOfCols).getValues()[0];
  }
  
  // returns data range in sheet, excluding  header row
  
  this.dstSheetGetDataRange = function(dstSheet) {
    
    var dstNumCols = dstSheet.getDataRange().getNumColumns();
    var dstNumRows = dstSheet.getDataRange().getNumRows();
    if (dstNumRows > 1) {
      return dstSheet.getRange(2, 1, dstNumRows-1, dstNumCols);  
    }
    return undefined;
  }
  
  // create an array with the column labels that exist in both the src and dst sheets.
  //
  // The keyLabel (defined as the  left-most label in the dst sheet) must also exist
  // in the src sheet.
  // Include column index in src and dst sheet.
  
  this.dstHeader2columns = function(srcHeader, dstSheet)  {
    
    var columns = {};
    var dstHeader = this.dstSheetGetHeaderValues(dstSheet);
    
    dstHeader.forEach(function(lbl) {
      
      var srcColIdx = srcHeader.indexOf(lbl);
      var dstColIdx = dstHeader.indexOf(lbl);
      var sameLabelInSrcSheet = srcColIdx >= 0;
      
      if (sameLabelInSrcSheet) {
        columns[lbl] = {src: srcColIdx, dst: dstColIdx};
      } else if (lbl == dstHeader[0]) {
        throw "keyLabel \"" + lbl + "\" (dst top-left cell value) missing in src labels";
      }  
    });
    return columns;
  }
  
  // clear the columns in the destination sheet (except the key column) that we're about to updated
  // (so later, empty fields indicate: a row that is no longer used)
  
  this.dstRangeClearColumns = function(columns, keyLabel, dstRange) {
    
    if (dstRange == undefined) {
      return;
    }
    for (cc in columns) {
      var isColumnWithKey = cc == keyLabel;
      var dstColumnExists = columns[cc].dst >= 0;
      if (!isColumnWithKey && dstColumnExists) {
        var columnIdx = columns[cc].dst;
        var columnToClean = dstRange.offset(0, columnIdx, dstRange.getHeight(), 1);
        columnToClean.clearContent();
      }
    }
  }
  
  // try to find the key in the rows that we started with
  
  this.findKeyInDstRange = function(columns, keyLabel, dstRange, key) {
    
    if (dstRange == undefined) {
      return -1;
    }
    var dstValues = dstRange.getValues();  // copy values locally
    
    for (var idx = 0; idx < dstValues.length; idx++) {  // -1 because range doesn't include header row
      var keyIdx = columns[keyLabel].dst;
      if (dstValues[idx][keyIdx] == key) {
        return idx;
      }    
    }        
    return -1;
  }   
  
  // appends a row to the dst sheet
  // return updated range and idx for the new row
  
  this.dstSheetInsertRow = function(columns, dstSheet, keyLabel, key) {
    
    dstSheet.appendRow(["new row"]);
    
    var dstRange = this.dstSheetGetDataRange(dstSheet);
    var columnIdx = columns[keyLabel].dst +1;
    
    dstRange.getCell(dstRange.getHeight(), columnIdx).setValue(key);  // -1 because no header row, getCell is 1-based
    dstRowIdx = dstRange.getHeight() - 1;  //  idx is 0-based  
    return [dstRange, dstRowIdx];
  }
  
  // copy values from the srcRow to the corresponding line in the dstRange
  
  this.dstRangeCopyFromSrc = function(columns, keyLabel, srcRow, dstRange, dstRowIdx) {
    
    for (cc in columns) {
      var isColumnWithKey = cc == keyLabel;
      var dstColumnExists = columns[cc].dst >= 0;
      if (!isColumnWithKey && dstColumnExists) {
        var srcColIdx = columns[cc].src;
        var dstColIdx = columns[cc].dst;
        dstRange.getCell(dstRowIdx+1, dstColIdx+1).setValue(srcRow[srcColIdx]);  // getCell() is 1-based
      }
    }  
  }
  
  this.main = function(srcSpreadsheetName, srcSheetName, dstSheetName) {

    var srcSheet = this.openSheet(srcSheetName, 2, false, srcSpreadsheetName);
    var dstSheet = this.openSheet(dstSheetName, 1, true);
    var srcValues = this.srcSheetGetValues(srcSheet);  // 1st row of src (and dst) is a header
    var srcHeader = srcValues.shift();
    
    var columns = this.dstHeader2columns(srcHeader, dstSheet);
    var keyLabel = Object.keys(columns)[0];
    var dstRange = this.dstSheetGetDataRange(dstSheet);
    
    this.dstRangeClearColumns(columns, keyLabel, dstRange);
    
    // for each row in the source, find the row with the corresponding key in the destination, 
    // then update the relevant columns in the destination
    
    var that = this;  // forEach has its own "this"
    srcValues.forEach(function(srcRow) {
      
      const key = srcRow[columns[keyLabel].src];
      
      var dstRowIdx = that.findKeyInDstRange(columns, keyLabel, dstRange, key);
      if (dstRowIdx < 0) {
        [dstRange, dstRowIdx] = that.dstSheetInsertRow(columns, dstSheet, keyLabel, key);
      }    
      that.dstRangeCopyFromSrc(columns, keyLabel, srcRow, dstRange, dstRowIdx)
    });  
  }
  
}).apply(OnMerge);

function onMerge() {
  var srcSpreadsheetName = "master";      // go-persons
  var srcSheetName = srcSpreadsheetName;  // tab has same name as workbook
  var dstSheetName = "test";
  
  OnMerge.main(srcSpreadsheetName, srcSheetName, dstSheetName);
}
