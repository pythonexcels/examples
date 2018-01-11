#
# Using ranges and offsets
#
# Jan 10, 2018 : Script modified to address problem
# with ws.Range("A6:B7,A9:B10").Value
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")
ws.Cells(1,1).Value = "Cell A1"
ws.Cells(1,1).Offset(2,4).Value = "Cell D2"
ws.Range("A2").Value = "Cell A2"
ws.Range("A3:B4").Value = "A3:B4"
try:
    ws.Range("A6:B7,A9:B10").Value = "A6:B7,A9:B10"
except:
    ws.Range("A6:B7;A9:B10").Value = "A6:B7,A9:B10"
wb.SaveAs('ranges_and_offsets.xlsx')
excel.Application.Quit()
