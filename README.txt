Thanks for taking the time to look into this code, and I hope you find it somewhat useful. 

Please leave any feedback on the Github main page (https://github.com/NixonInnes/ExcelVBA-RecordCellChange), or alternatively email me at nixoninnes AT gmail DOT com.

This was written to track changes made to an excel spreadsheet. There is an in-built functionality to do this; however I found it to be rather clunky and slow.

Changes made to a ROW in excel will be recorded in their respective COLUMS A & B. The date of the change is recorded into column A; the old and new values of the cells are recorded into column B.

PREREQUISITES
==============
- Microsoft Scripting Runtime
	- Open Excel
	- Navigate to the VBA Editor (Alt+F11)
	- Tools -> References
	- Check Microsoft Scripting Runtime

ADDING THE .CLS
================
- Open Excel
- Navigate to the VBA Editor (Alt+F11)
- File -> Import
- Select the downloaded .cls file
