# Google Sheets Macros

*If you enjoy my [Google Workspace Apps Script](https://developers.google.com/apps-script) work, please consider buying me a cup of coffee!* 


[![](https://techstreams.github.io/images/bmac.svg)](https://www.buymeacoffee.com/techstreams)

---

This repository contains custom macros for Google Sheets.

*(For more information on Google Sheets macros, see [Automate tasks in Google Sheets](https://support.google.com/docs/answer/7665004)).*

---

## Installation

**To manually install a Google Sheets macro:**

1. Open a new Google Sheet at [sheets.google.com](http://sheets.google.com/)
1. Open the script editor by selecting the **Tools** > **Script Editor** menu
1. Copy and paste the desired macro code to the bottom of the **Code.gs** *(or open script)* file
1. Click the **Save** icon or select the **File** > **Save** menu option to save the macro
1. Close the Script Editor window

**To import a Google Sheets macro:**

1. Select the Google Sheets **Tools** > **Macros** > **Import** menu
1. Identify the macro to import and click the associated **Add function** option

**To run a Google Sheets macro:**

1. Select the Google Sheets **Tools** > **Macros** > *'name of macro'* menu
  

---

## Macros

* **Benford's Law**  
  * [Macro Code](/BenfordsLaw.gs)
  * [Blog Post](https://medium.com/@techstreams/using-apps-script-a-google-sheets-macro-and-benfords-law-to-detect-potential-fraud-9fbd91b325ab)
  
* **Dates**
  * [AddCalendarDropdown()](/Dates.gs) - Add calendar dropdown and date validation to active range.
  
* **Freeze**
  * [FreezeActiveSheet()](/Freeze.gs) - Freeze specified rows and columns on active sheet
  * [FreezeAllSheets()](/Freeze.gs) - Freeze specified rows and columns on all sheets

* **TSCreateUrlCheatsheet**
  * [TSCreateUrlCheatsheet()](/TSCreateUrlCheatsheet.gs) - Create Google Workspace New URLs Cheatsheet
  * *Aware of any other Google Workspace new resource URLs?  Please let me know through the [issues](https://github.com/techstreams/google-sheets-macros/issues).*
  
---

## License

**google-sheets-macros License**

© Laura Taylor ([github.com/techstreams](https://github.com/techstreams)). Licensed under an MIT license.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
