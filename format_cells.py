#
# Format cell font name and size, format numbers in monetary format
#
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")

for i,font in enumerate(["Arial","Courier New","Garamond","Georgia","Verdana"]):
    ws.Range(ws.Cells(i+1,1),ws.Cells(i+1,2)).Value = [font,i+i]
    ws.Range(ws.Cells(i+1,1),ws.Cells(i+1,2)).Font.Name = font
    ws.Range(ws.Cells(i+1,1),ws.Cells(i+1,2)).Font.Size = 12+i
    
ws.Range("A1:A5").HorizontalAlignment = win32.constants.xlRight
ws.Range("B1:B5").NumberFormat = "$###,##0.00"
ws.Columns.AutoFit()
wb.SaveAs('format_cells.xlsx')
excel.Application.Quit()
