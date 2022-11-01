#
# make15x15.py
# Expand an existing 10x10 multiplication table and resize columns
#
import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open('MultiplicationTable.xlsx')
excel.Visible = True
ws = wb.Worksheets('Sheet1')
ws.Range("B11:K11").Select()
excel.Selection.AutoFill(ws.Range("B11:K16"), win32.constants.xlFillDefault)
ws.Range("K2:K16").Select()
excel.Selection.AutoFill(ws.Range("K2:P16"), win32.constants.xlFillDefault)
ws.Columns("B:P").Select()
excel.Selection.ColumnWidth = 4
wb.SaveAs('NewMultiplicationTable.xlsx')
excel.Application.Quit()
