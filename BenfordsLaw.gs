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
 * Google Sheets macro for automating data analysis utilizing Benford's Law
 * See Blog Post - https://medium.com/@techstreams/using-apps-script-a-google-sheets-macro-and-benfords-law-to-detect-potential-fraud-9fbd91b325ab
 */
function BenfordsLaw() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      activeSheet = ss.getActiveSheet(),
      dataRange = activeSheet.getDataRange(),
      numBenfordCols = 5,
      analyze, analysisColumn, benfords, benfordsRange, button, chart, chartRange, count, countRange, formulas, frequency, 
      frequencyRange, headers, response, sample, sampleRange, start, startColumn, sum, text, xAxisLabelRange, ui;
  ui = SpreadsheetApp.getUi();
  response = ui.prompt('Column to Analyze', 'Enter Letter of Column of Data to Analyze (ex. A)', ui.ButtonSet.OK_CANCEL);
  button = response.getSelectedButton();
  text = response.getResponseText();
  if (button == ui.Button.OK) {
    if (text !== "" && /^[a-zA-Z]+$/.test(text)) {
      analyze = text.toUpperCase();
      response = ui.prompt("Benford's Law Start Column", "Enter Letter of Column to Begin Benford's Law Calculations (ex. D)", Browser.Buttons.OK_CANCEL);
      button = response.getSelectedButton();
      text = response.getResponseText();
      if (button == ui.Button.OK) {
        if (text !== "" && /^[a-zA-Z]+$/.test(text)) {
          start = text.toUpperCase();
          analysisColumn = activeSheet.getRange(analyze+':'+analyze);
          startColumn = activeSheet.getRange(start+':'+start);     
          analysisColumn.setNumberFormat('General');
          activeSheet.insertColumnsAfter(startColumn.getColumn(), numBenfordCols-1);
          headers = [['First Digit','Count','Frequency','Benford Rate','Sample Rate']];
          activeSheet.getRange(1, startColumn.getColumn(), 1, 5).setValues(headers).setFontWeight('bold').setHorizontalAlignment('center');
          activeSheet.autoResizeColumns(startColumn.getColumn(), numBenfordCols);
          formulas = [];
          for (var i=2; i<=dataRange.getLastRow(); i++) {
            var f = ['=LEFT('+analyze+i+',1)'];
            formulas.push(f);    
          }
          activeSheet.getRange(2, startColumn.getColumn(), dataRange.getLastRow()-1).setFormulas(formulas);
          countRange = activeSheet.getRange(2, startColumn.getColumn()+1, 9, 1);
          count = activeSheet.getRange(2, startColumn.getColumn()+1);
          count.setValue('1').setNumberFormat('@').autoFill(countRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
          frequencyRange = activeSheet.getRange(2, startColumn.getColumn()+2, 9, 1);
          frequency = activeSheet.getRange(2, startColumn.getColumn()+2);
          frequency.setFormula('=COUNTIF('+start+':'+start+','+count.getA1Notation()+')').autoFill(frequencyRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
          sum = activeSheet.getRange(11, startColumn.getColumn()+2)
          sum.setFormula('=SUM('+frequencyRange.getA1Notation()+')').setFontWeight('bold');
          benfordsRange = activeSheet.getRange(2, startColumn.getColumn()+3, 9, 1);
          benfords = activeSheet.getRange(2, startColumn.getColumn()+3);
          benfords.setFormula('=LOG10(1/'+count.getA1Notation()+'+1)').setNumberFormat('0.00%').autoFill(benfordsRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
          sampleRange = activeSheet.getRange(2, startColumn.getColumn()+4, 9, 1);
          sample = activeSheet.getRange(2, startColumn.getColumn()+4);
          sample.setFormula('='+frequency.getA1Notation()+'/INDIRECT(ADDRESS(' + sum.getRow() + ',' + sum.getColumn() + '))').setNumberFormat('0.00%');
          sample.autoFill(sampleRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
          xAxisLabelRange = activeSheet.getRange(1, startColumn.getColumn()+1, 10, 1);
          chartRange = activeSheet.getRange(1, startColumn.getColumn()+3, 10, 2);
          chart = activeSheet.newChart()
                             .asLineChart()
                             .setTitle("Benford's Law Analysis")
                             .setOption('legend.textStyle.fontSize', 14)
                             .addRange(xAxisLabelRange)
                             .addRange(chartRange)
                             .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
                             .setTransposeRowsAndColumns(false)
                             .setNumHeaders(-1)
                             .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
                             .setOption('useFirstColumnAsDomain', true)
                             .setOption('curveType', 'none')
                             .setOption('domainAxis.direction', 1)
                             .setOption('series.0.labelInLegend', 'Benford Rate')
                             .setOption('series.0.lineWidth', 4)
                             .setOption('series.1.labelInLegend', 'Sample Rate')
                             .setOption('series.1.lineWidth', 4)
                             .setPosition(3, startColumn.getColumn()+1, 14, 78)
                             .build();
          activeSheet.insertChart(chart);
        } else {
          ui.alert("ERROR: Invalid Value", "Please re-run Benford's Law macro and enter a valid column letter for start of Benford's Law calculations (ex. D)" , Browser.Buttons.OK)
        }
      } else {
        ui.alert("Benford's Law macro has been canceled.");
      }
    } else {
      ui.alert("ERROR: Invalid Value", "Please re-run Benford's Law macro and enter a valid column letter for data to be analyzed (ex. A)" , Browser.Buttons.OK)
    }
  } else {
    ui.alert("Benford's Law macro has been canceled.");
  }
}


