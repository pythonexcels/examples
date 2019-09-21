#
# driving.py
#
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True
wb = excel.Workbooks.Add()
ws = wb.Worksheets('Sheet1')
ws.Name = 'Built with Python'
ws.Cells(1,1).Value = 'Hello Excel'
print (ws.Cells(1,1).Value)
for i in range(1,5):
    ws.Cells(2,i).Value = i  # Don't do this
ws.Range(ws.Cells(3,1),ws.Cells(3,4)).Value = [5,6,7,8]
ws.Range("A4:D4").Value = [i for i in range(9,13)]
ws.Cells(5,4).Formula = '=SUM(A2:D4)'
ws.Cells(5,4).Font.Size = 16
ws.Cells(5,4).Font.Bold = True


