################
Project: Save As OFX file from Excel
Purpose: To convert imple Data in Excel sheet to either OFX for import into account software

Version: 0.2
Date written: 27 January 2019

Author: Craig Lambie
Author URI: craiglambie.com

Data:
3 Columns: Date, Description or Memo and Amount

Routines: 
SaveAsOFX - saves the current active sheet as a OFX file with the same name as the open file
Amended version of Craig Lambie's routine which is based upon the XLS2OFX Converter v1.0 by Josep Bori

Usage:
1. Open your Excel file (of any type)
2. Push the data you want to a clean sheet with 3 headers/ top row
3. Go to ribbon View>Macros
4. On dialogue box, dropdown box "Macros in:" select "All Workbooks"
5. Select SaveAsOFX as required
6. Visit the location your excel file is saved to find the OFX file

Installation:
1. Open Excel
2. Open VBA (Alt + F11)
3. Import the .cls file
