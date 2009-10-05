#
# Copy data and formatting from a range of one worksheet
# to all other worksheets in a workbook
#
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")
ws.Range("A1:J10").Formula = "=row()*column()"
wb.Worksheets.FillAcrossSheets(wb.Worksheets("Sheet1").Range("A1:J10"))
wb.SaveAs('copy_worksheet_to_worksheet.xlsx')
excel.Application.Quit()
