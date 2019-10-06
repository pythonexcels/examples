# Example Files from Pythonexcels.com

This repository contains example Python scripts and Excel files described in the blog [https://www.pythonexcels.com/](https://www.pythonexcels.com).

## ABCDCatering.xls

This Excel file contains the sample spreadsheet used in many of the pivot table examples in this repository. ABCDCatering.xls is described in [Cleaning Up Corporate ERP Data](https://pythonexcels.com/python/2009/11/09/Cleaning-Up-Corporate-ERP-Data.html).

## add_a_workbook.py

This script starts Excel, adds a workbook, and saves the empty workbook.  This script is described in [Python Excel Mini Cookbook](https://pythonexcels.com/python/2009/10/05/Python-Excel-Mini-Cookbook.html).

## add_a_worksheet.py

This script creates a new Excel workbook with three sheets, adds a fourth worksheet, names it MyNewSheet, and saves the workbook to the file add_a_worksheet.xlsx. This script is described in  [Python Excel Mini Cookbook](https://pythonexcels.com/python/2009/10/05/Python-Excel-Mini-Cookbook.html).

## autofill_cells.py

This script uses Excels autofill capability to examine data in cells A1 and A2, then autofill the remaining cells through A10.  Excel spreadsheet is written to autofill_cells.xlsx.  This script is described in [Python Excel Mini Cookbook](https://pythonexcels.com/python/2009/10/05/Python-Excel-Mini-Cookbook.html).

## cell_color.py

This script illustrates adding an interior color to the cell using `Interior.ColorIndex`.  Column A, rows 1 through 20 are filled with a number and assigned that ColorIndex.  The spreadsheet is written to cell_color.xlsx.  This script is described in  [Python Excel Mini Cookbook](https://pythonexcels.com/python/2009/10/05/Python-Excel-Mini-Cookbook.html).

## column_widths.py

This script creates two columns of data, one narrow and one wide, then formats the column width with the `ColumnWidth` property.  You can also use the `Columns.AutoFit()` function to autofit all columns in the spreadsheet.  The spreadsheet is written to column_widths.xlsx.  This script is described in  [Python Excel Mini Cookbook](https://pythonexcels.com/python/2009/10/05/Python-Excel-Mini-Cookbook.html).

## conditionalformatting.py

This script builds two data tables from scratch, applies conditional formatting to the tables, and saves the result to ConditionalFormatting.xlsx. This script is described in [Mapping Excel VB Macros to Python Revisited](https://pythonexcels.com/python/2009/10/20/Mapping-Excel-VB-Macros-to-Python-Revisited.html).

## copy_worksheet_to_worksheet.py

This script uses `FillAcrossSheets()` to copy data from one location to all other worksheets in the workbook.  Specifically, the data in the range A1:J10 is copied from Sheet1 to Sheet2 and Sheet3.  The spreadsheet is written to copy_worksheet_to_worksheet.xlsx.  This script is described in  [Python Excel Mini Cookbook](https://pythonexcels.com/python/2009/10/05/Python-Excel-Mini-Cookbook.html).

## driving.py

This script provides a simple introduction to opening Excel by creating a workbook, creating a worksheet, and adding some data to the worksheet.  This script is best run by entering the text line-by-line into Python. This script is described in
[Basic Excel Driving With Python](https://pythonexcels.com/python/2009/09/29/Basic-Excel-Driving-With-Python.html).

## erpdata.py

This script loads the spreadsheet file ABCDCatering.xls, prepares it for pivot table insertion and saves the file.  The output spreadsheet is written to the file newABCDCatering.xls. This script is described in [Cleaning Up Corporate ERP Data](https://pythonexcels.com/python/2009/11/09/Cleaning-Up-Corporate-ERP-Data.html).

## erpdatapivot.py

This script extends the erpdata.py script by building 5 pivot tables based on the input spreadsheet file ABCDCatering.xls.  The output spreadsheet is written to the file newABCDCatering.xls.  This script is described in [Automating Pivot Tables with Python](https://pythonexcels.com/python/2009/11/23/Automating-Pivot-Tables-with-Python.html).

## erppivotdragdrop.py

erppivotdragdrop.py is based on erppivotextended.py and provides a simple user interface for running the script. You can drag and drop multiple files onto the script; when complete, the script issues a simple message box telling you when everything is done. The script prepares the poorly formatted table data table from ABCDCatering.xls for pivot table insertion, inserts additional data columns derived from the existing data, and creates six pivot tables.  The output spreadsheet is written to ABCDCatering_new.xlsx.  This script is described in [A User Friendly Experience](https://pythonexcels.com/python/2010/02/07/A-User-Friendly-Experience.html).

## erppivotextended.py

This script is based on erpdatapivot.py and adds column insertion to derive new data columns for extended data analysis.  The script prepares the poorly formatted table data table from ABCDCatering.xls for pivot table insertion, inserts additional data columns derived from the existing data, and creates six pivot tables.  The output spreadsheet is written to the file newABCDCatering.xls. This script is described at [Extending Pivot Table Data](https://pythonexcels.com/python/2009/12/03/Extending-Pivot-Table-Data.html).

## format_cells.P

This script creates two columns of data, then formats the font type and font size used in the worksheet.  Five different fonts and sizes are used, the numbers are formatted using a monetary format.  The spreadsheet is written to format_cells.xlsx.  This script is described in  [Python Excel Mini Cookbook](https://pythonexcels.com/python/2009/10/05/Python-Excel-Mini-Cookbook.html).

## make15x15.py

This script loads the file My Documents\MultiplicationTable.xlsx, expands the multiplication table from 10x10 to 15x15, changes the column width, and saves the updated worksheet to My Documents\NewMultiplicationTable.xlsx.  This script is described in [Mapping Excel VB Macros to Python](https://pythonexcels.com/python/2009/10/12/Mapping-Excel-VB-Macros-to-Python.html).

## MultiplicationTable.xlsx

Simple 10x10 multiplication source file for make15x15.py.  This script is described in  [Mapping Excel VB Macros to Python](https://pythonexcels.com/python/2009/10/12/Mapping-Excel-VB-Macros-to-Python.html).

## newABCDCatering.xls

newABCDCatering.xls is the Excel spreadsheet output from erpdata.py and contains a well formatted data table for pivot table conversion.  This file is described in the [Introducing Pivot Tables](https://pythonexcels.com/python/2009/11/11/Introducing-Pivot-Tables.html).

## open_an_existing_workbook.py

This script opens an existing workbook and displays it (note the statement `excel.Visible = True`).  The  workbook1.xlsx file must exist in your  “My Documents” directory.  You can also open spreadsheet files by specifying the full path to the file as shown below.  Using r' in the statement r'C:\myfiles\excel\workbook2.xlsx' automatically escapes the backslash characters and makes the file name a bit more concise.  This script is described in  [Python Excel Mini Cookbook](https://pythonexcels.com/python/2009/10/05/Python-Excel-Mini-Cookbook.html).

## payrates.py

This script opens each spreadsheet file in the current directory, looks for specific information in the script, and writes a CSV file with the information. Sample data for this script is available in the Payroll folder. This script is described in [Ninety Six Spreadsheets](https://pythonexcels.com/python/2012/09/22/Ninety-Six-Spreadsheets.html).

## ranges_and_offsets.py

This script uses some different techniques for addressing cells using the `Cells()` and `Range()` operators.   This script is described in  [Python Excel Mini Cookbook](https://pythonexcels.com/python/2009/10/05/Python-Excel-Mini-Cookbook.html).

## row_height.py

Similar to column height, row height can be set with the RowHeight setting.  You can also use `AutoFit()` to automatically adjust the row height based on cell contents.  This script is described in  [Python Excel Mini Cookbook](https://pythonexcels.com/python/2009/10/05/Python-Excel-Mini-Cookbook.html).

<em>See [Python Excels](https://pythonexcels.com) for more information on these scripts.</em>
