#
# Open an existing workbook
#
import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open('workbook1.xlsx')
# Alternately, specify the full path to the workbook
# wb = excel.Workbooks.Open(r'C:\myfiles\excel\workbook2.xlsx')
excel.Visible = True
