#
# Add a workbook, add a worksheet,
# name it 'MyNewSheet' and save
#
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Add()
ws = wb.Worksheets.Add()
ws.Name = "MyNewSheet"
wb.SaveAs('add_a_worksheet.xlsx')
excel.Application.Quit()
