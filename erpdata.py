#
# erpdata.py: Load raw EPR data and clean up header info
#
import win32com.client as win32
import sys
excel = win32.gencache.EnsureDispatch('Excel.Application')
#excel.Visible = True
try:
    wb = excel.Workbooks.Open('ABCDCatering.xls')
except:
    print "Failed to open spreadsheet ABCDCatering.xls"
    sys.exit(1)
ws = wb.Sheets('Sheet1')
xldata = ws.UsedRange.Value
newdata = []
for row in xldata:
    if len(row) == 13 and row[-1] is not None:
        newdata.append(list(row))
lasthdr = "Col A"
for i,field in enumerate(newdata[0]):
    if field is None:
        newdata[0][i] = lasthdr + " Name"
    else:
        lasthdr = newdata[0][i]
wsnew = wb.Sheets.Add()
wsnew.Range(wsnew.Cells(1,1),wsnew.Cells(len(newdata),len(newdata[0]))).Value = newdata
wsnew.Columns.AutoFit()
if int(float(excel.Version)) >= 12:
    wb.SaveAs('newABCDCatering.xlsx',win32.constants.xlOpenXMLWorkbook)
else:
    wb.SaveAs('newABCDCatering.xls')
excel.Application.Quit()
