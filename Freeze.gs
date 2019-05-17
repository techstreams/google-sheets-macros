/*
* Copyright Laura Taylor (https://github.com/techstreams/google-sheets-macros)
*
* Permission is hereby granted, free of charge, to any person obtaining a copy
* of this software and associated documentation files (the "Software"), to deal
* in the Software without restriction, including without limitation the rights
* to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
* copies of the Software, and to permit persons to whom the Software is
* furnished to do so, subject to the following conditions:
*
* The above copyright notice and this permission notice shall be included in all
* copies or substantial portions of the Software.
*
* THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
* IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
* FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
* AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
* LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
* OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
* SOFTWARE.
*/

/** @OnlyCurrentDoc */

/**
 * Google Sheets macro for freezing rows and columns on active sheet
 */
function FreezeActiveSheet() {
  var ss = SpreadsheetApp.getActive(),
      sheet = ss.getActiveSheet(),
      ui = SpreadsheetApp.getUi(),
      button, columns, response, rows, text;
  response = ui.prompt('Freeze to Row', 'Enter Number of Rows to Freeze', ui.ButtonSet.OK_CANCEL);
  button = response.getSelectedButton();
  text = response.getResponseText();
  if (button == ui.Button.OK) {
    if (text !== "" && /^[0-9]+$/.test(text)) {
      rows = parseInt(text, 10);
      response = ui.prompt('Freeze to Column', 'Enter Number of Columns to Freeze', ui.ButtonSet.OK_CANCEL);
      button = response.getSelectedButton();
      text = response.getResponseText();
      if (button == ui.Button.OK) {
        if (text !== "" && /^[0-9]+$/.test(text)) {
          columns = parseInt(text, 10);
          sheet.setFrozenRows(rows);
          sheet.setFrozenColumns(columns);
          ui.alert("Macro complete.");
        } else {
          ui.alert("Invalid column value. Macro has been cancelled.");
        }
      }  else {
        ui.alert("Macro has been canceled.");
      }
    } else {
      ui.alert("Invalid row value. Macro has been cancelled.");
    }
  } else {
    ui.alert("Macro has been canceled.");
  }
};

/**
 * Google Sheets macro for freezing rows and columns on all sheets
 */
function FreezeAllSheets() {
  var ss = SpreadsheetApp.getActive(),
      sheets = ss.getSheets(),
      ui = SpreadsheetApp.getUi(),
      button, columns, response, rows, text;
  response = ui.prompt('Freeze to Row', 'Enter Number of Rows to Freeze', ui.ButtonSet.OK_CANCEL);
  button = response.getSelectedButton();
  text = response.getResponseText();
  if (button == ui.Button.OK) {
    if (text !== "" && /^[0-9]+$/.test(text)) {
      rows = parseInt(text, 10);
      response = ui.prompt('Freeze to Column', 'Enter Number of Columns to Freeze', ui.ButtonSet.OK_CANCEL);
      button = response.getSelectedButton();
      text = response.getResponseText();
      if (button == ui.Button.OK) {
        if (text !== "" && /^[0-9]+$/.test(text)) {
          columns = parseInt(text, 10);
          sheets.forEach(function(sheet) {
            sheet.setFrozenRows(rows);
            sheet.setFrozenColumns(columns);
          });
          ui.alert("Macro complete.");
        } else {
          ui.alert("Invalid column value. Macro has been cancelled.");
        }
      }  else {
        ui.alert("Macro has been canceled.");
      }
    } else {
      ui.alert("Invalid row value. Macro has been cancelled.");
    }
  } else {
    ui.alert("Macro has been canceled.");
  }
};
