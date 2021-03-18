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
 * Google Sheets macro for creating a Google Workspace new URLs cheatsheet
 */
function TSCreateUrlCheatsheet() {
  const ss = SpreadsheetApp.getActive(),
        sheet = ss.getActiveSheet()
        titles = [['DOCS','SHEETS','SLIDES','FORMS','KEEP','CALENDAR','SITES','SCRIPT']],
        data = [['doc.new','sheet.new','slide.new','form.new','keep.new','cal.new','site.new','script.new'],
                ['docs.new','sheets.new','slides.new','forms.new','note.new','meeting.new','sites.new',''],
                ['document.new','spreadsheet.new','deck.new','','notes.new','','website.new',''],
                ['','','presentation.new','','','','','']];
  let banding;

  sheet.getRange('A1:H1').setValues(titles).setHorizontalAlignment('center')
       .setFontWeight('body')
       .setFontSize(14)
       .setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID)
       .activate();
  sheet.getRange('A2:H5').setValues(data).setFontSize(12).activate();
  sheet.getRange('A1:H5').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  sheet.getRange('A1:H5').getBandings()[0].setHeaderRowColor('#8bc34a').setFirstRowColor('#ffffff')
                         .setSecondRowColor('#eef7e3').setFooterRowColor(null);
  sheet.getRange('A5:H5')
       .setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  sheet.autoResizeColumns(1, 8);
  sheet.getRange('A7').setValue('For more macros by @techstreams see:').setFontSize(10).activate();
  sheet.getRange('C7').setValue('https://github.com/techstreams/google-sheets-macros')
       .setFontSize(10).activate();
};
