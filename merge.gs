"use strict";
/*                                    MERGE GOOGLE SHEETS
*
*                                         Coert Vonk
*                                        February 2019
*
*                         https://github.com/cvonk/gas-sheets-merge
* USE:
*
* The names of the src spreadsheet, src sheet and dst sheet are at the end of this code.
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
* refer to https://github.com/cvonk/gas-sheets-merge/blob/master/README.md
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

// update dst sheet based on src sheet

function onMerge(parameters) {
  
  // create an array with the column labels that exist in both the src and dst sheets.
  //
  // The keyLabel (defined as the  left-most label in the dst sheet) must also exist
  // in the src sheet.
  // Include column index in src and dst sheet.
  
  function _dstHeader2columns(srcHeader, dstHeader)  {
    
    var columns = {};
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
  
  function _dstRangeClearColumns(columns, keyLabel, dstRange) {
    
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
  
  function _findKeyInDstRange(columns, keyLabel, dstRange, key) {
    
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
  
  function _dstSheetInsertRow(columns, dstSheet, keyLabel, key) {
    
    dstSheet.appendRow(["new row"]);
    
    var dstRange = Common.sheetGetDataRange(dstSheet);
    var columnIdx = columns[keyLabel].dst +1;
    
    dstRange.getCell(dstRange.getHeight(), columnIdx).setValue(key);  // -1 because no header row, getCell is 1-based
    dstRowIdx = dstRange.getHeight() - 1;  //  idx is 0-based  
    return [dstRange, dstRowIdx];
  }
  
  // copy values from the srcRow to the corresponding line in the dstRange
  
  function _dstRangeCopyFromSrc(columns, keyLabel, srcRow, dstRange, dstRowIdx) {
    
    for (cc in columns) {
      var isColumnWithKey = cc == keyLabel;
      var dstColumnExists = columns[cc].dst >= 0;
      if (!isColumnWithKey && dstColumnExists) {
        var srcColIdx = columns[cc].src;
        var dstColIdx = columns[cc].dst;
        dstRange.getCell(dstRowIdx + 1, dstColIdx + 1).setValue(srcRow[srcColIdx]);  // getCell() is 1-based
      }
    }  
  }

  // validate parameters
  
  if (parameters.srcSpreadsheetId == undefined ||
      parameters.srcSheetName == undefined ||
      parameters.dstSheetName == undefined) {
    throw("invalid parameters to onMerge()");
  }
  if (typeof parameters.srcSpreadsheetId !== "string" ||
      typeof parameters.srcSheetName !== "string" ||
      typeof parameters.dstSheetName !== "string" ) {
    throw("invalid parameters type to onMerge()");
  }
  
  var srcSpreadsheet = Common.spreadsheetOpenById(parameters.srcSpreadsheetId);
  var srcSheet = Common.sheetOpen(srcSpreadsheet, parameters.srcSheetName, 2, true );
  var srcValues = srcSheet.getDataRange().getValues();
  var srcHeader = srcValues.shift();
  
  var dstSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dstSheet = Common.sheetOpen(dstSpreadsheet, parameters.dstSheetName, 1, true);
  var dstHeader = Common.sheetGetHeaderRange(dstSheet).getValues()[0];    
  var columns = _dstHeader2columns(srcHeader, dstHeader);
  var keyLabel = Object.keys(columns)[0];
  var dstRange = Common.sheetGetDataRange(dstSheet);
  
  _dstRangeClearColumns(columns, keyLabel, dstRange);
  
  // for each row in the source, find the row with the corresponding key in the destination, 
  // then update the relevant columns in the destination
  
  srcValues.forEach(function(srcRow) {
    
    var key = srcRow[columns[keyLabel].src];	
    var dstRowIdx = _findKeyInDstRange(columns, keyLabel, dstRange, key);
    if (dstRowIdx < 0) {
      [dstRange, dstRowIdx] = _dstSheetInsertRow(columns, dstSheet, keyLabel, key);
    }    
    _dstRangeCopyFromSrc(columns, keyLabel, srcRow, dstRange, dstRowIdx)
  });  
}

function onMerge_dbg() {
  onMerge({srcSpreadsheetId: "YourSpreadSheetId", // eg output from LDAP
           srcSheetName: "go-persons", 
           dstSheetName: "persons"});    
}
