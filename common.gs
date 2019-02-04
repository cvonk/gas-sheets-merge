"use strict";
/*                             FUNCTIONS COMMON TO VARIOUS SCRIPT
*
*                                         Coert Vonk
*                                        February 2019
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

// create pivot sample data from user allocations

var Common = {};
(function() {
  
  this.spreadsheetOpenByName = function(spreadsheetName) {

    var files = DriveApp.searchFiles('mimeType = "' + MimeType.GOOGLE_SHEETS + '" and title contains "' + spreadsheetName + '"');
    while (files.hasNext()) {
      var spreadsheet = SpreadsheetApp.open(files.next());
      if (spreadsheet == undefined) {
        throw "Spreadsheet \"" + spreadsheetName + "\" not found on gDrive";
      }
      return spreadsheet;  
    }
    return undefined;
  }

  this.spreadsheetOpenById = function(spreadsheetId) {

    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    if (spreadsheet == undefined) {
        throw "Spreadsheet \"" + spreadsheetId + "\" not found";
      return undefined;
    }
    return spreadsheet;  
  }
  
  this.sheetCreate = function(spreadsheet, sheetName, overwriteSheet) {

    var sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet == undefined) {
      sheet = spreadsheet.insertSheet(sheetName);
      if (sheet == undefined) {
        throw "Can't create sheet \"" + sheetName + "\" in spreadsheet \"" + spreadsheet.getName() + "\"";
      }
    } else {
      if (overwriteSheet) {
        sheet.clear();
      } else {
        throw "Sheet \"" + sheetName + "\" already exists in spreadsheet \"" + spreadsheet.getName() + "\"";
      }
      
    } 
    return sheet;
  }
  
  this.sheetOpen = function(spreadsheet, sheetName, minNrOfDataRows, requireLabelInA1) {

    var sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet == undefined) {
      throw "Sheet \"" + sheetName + "\" not found in spreadsheet \"" + spreadsheet.getName() + "\"";
    }
    if (requireLabelInA1 && sheet.getRange("A1").isBlank()) {
      throw "Sheet \"" + sheetName + "\" (part of " + spreadsheet.getName() + ") doesn't have a keyLabel in A1";
    }
    if (sheet.getDataRange().getNumRows() < minNrOfDataRows) {
      throw "Sheet \"" + sheetName + "\" (part of " + spreadsheet.getName() + ") doesn't have at least " + minNrOfDataRows + " rows";
    }
    return sheet;
  }
  
  // returns data range in sheet, excluding  header row
  
  this.sheetGetHeaderRange = function(sheet) {
    
    var numCols = sheet.getDataRange().getNumColumns();
    return sheet.getRange(1, 1, 1, numCols);  
  }
  
  // returns data range in sheet, excluding  header row
  
  this.sheetGetDataRange = function(sheet) {
    
    var numCols = sheet.getDataRange().getNumColumns();
    var numRows = sheet.getDataRange().getNumRows();
    if (numRows > 1) {
      return sheet.getRange(2, 1, numRows - 1, numCols);  
    }
    return undefined;
  }
    
  this.arrayHasNoEmptyEl = function(array) {

    return array.indexOf("") == -1;
  }
  
  this.getFilteredDataRange = function(sheet) {
    
    var all = sheet.getDataRange().getValues();
    var filtered = [];
    
    for (var ii = 0; ii < all.length; ii++ ) {
      if (!sheet.isRowHiddenByFilter(ii+1)) {
        filtered.push(all[ii]);
      }
    }
    return filtered;
  }
  
}).apply(Common);
