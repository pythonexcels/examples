#
# Set column widths
#
import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")
ws.Range("A1:A10").Value = "A"
ws.Range("B1:B10").Value = "This is a very long line of text"
ws.Columns(1).ColumnWidth = 1
ws.Range("B:B").ColumnWidth = 27
# Alternately, you can autofit all columns in the worksheet
# ws.Columns.AutoFit()
wb.SaveAs('column_widths.xlsx')
excel.Application.Quit()
