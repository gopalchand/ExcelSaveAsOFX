################
Project: Save As OFX file from Excel
Purpose: To convert simple Data in Excel sheet to  OFX for import into accounting software

Version: 0.2
Date modified: 13 January 2019

Author: Gopal Chand

Data:
3 Columns: Date, Description/Memo and Amount

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

Known issues and workarounds:
The presence of SGML predefined entities (", <, >, &, ') in the Description/Memo field is likely to lead to load failures

For credit card accounts use the following:
sTranAmount = -1# * rgeTransactionList.Offset(iTransaction, 2).Value

If a separate memo field is available as column 3 and the amount as column 4 then use         
sTranMemo = rgeTransactionList.Offset(iTransaction, 2).Value
sTranAmount = rgeTransactionList.Offset(iTransaction, 3).Value
