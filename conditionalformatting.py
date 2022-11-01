#
# conditionalformatting.py
# Create two tables and apply conditional formatting
#
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
# excel.Visible = True
wb = excel.Workbooks.Add()
ws = wb.Worksheets('Sheet1')
ws.Range("B2:K2").Value = [i for i in range(1, 11)]
ws.Range("B2:B11").Value = list(zip([i for i in range(1, 11)]))
ws.Range("C3").Formula = "=$B3*C$2"
ws.Range("C3:C3").Select()
excel.Selection.AutoFill(ws.Range("C3:K3"), win32.constants.xlFillDefault)
ws.Range("C3:K3").Select()
excel.Selection.AutoFill(ws.Range("C3:K11"), win32.constants.xlFillDefault)
# Add the table of random integers
ws.Range("B13:K22").Formula = "=INT(RAND()*100)"
ws.Range("B2:K22").Select()
excel.Selection.FormatConditions.AddColorScale(ColorScaleType=3)
excel.Selection.FormatConditions(excel.Selection.FormatConditions.Count).SetFirstPriority()
[csc1, csc2, csc3] = [excel.Selection.FormatConditions(1).ColorScaleCriteria(n) for n in range(1, 4)]
csc1.Type = win32.constants.xlConditionValueLowestValue
csc1.FormatColor.Color = 13011546
csc1.FormatColor.TintAndShade = 0
csc2.Type = win32.constants.xlConditionValuePercentile
csc2.Value = 50
csc2.FormatColor.Color = 8711167
csc2.FormatColor.TintAndShade = 0
csc3.Type = win32.constants.xlConditionValueHighestValue
csc3.FormatColor.Color = 7039480
csc3.FormatColor.TintAndShade = 0
ws.Range("B:K").ColumnWidth = 4
ws.Range("A1").Select()
wb.SaveAs('ConditionalFormatting.xlsx')
excel.Application.Quit()
