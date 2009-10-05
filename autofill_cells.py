#
# Autofill cell contents
#
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")
ws.Range("A1").Value = 1
ws.Range("A2").Value = 2
ws.Range("A1:A2").AutoFill(ws.Range("A1:A10"),win32.constants.xlFillDefault)
wb.SaveAs('autofill_cells.xlsx')
excel.Application.Quit()
