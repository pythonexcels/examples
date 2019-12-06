#
# payrates.py
# Report payrates for two employees across multiple spreadsheets
#
import win32com.client as win32
import glob
import os

xlxsfiles = sorted(glob.glob("*.xlsx"))
print ("Reading %d files..."%len(xlxsfiles))

steve = []
jeff = []
cwd = os.getcwd()
excel = win32.gencache.EnsureDispatch('Excel.Application')
fpjeffsteve = open('jeffsteve.csv','w')
for xlsxfile in xlxsfiles:
    wb = excel.Workbooks.Open(cwd+"\\"+xlsxfile)
    try:
        ws = wb.Sheets('PAYROLL')
    except:
        print ("No worksheet named 'PAYROLL' in %s, skipping"%xlsxfile)
        jeff.append(0.0)
        steve.append(0.0)
        wb.Close()
        continue
    xldata = ws.UsedRange.Value
    names = [r[1] for r in xldata]
    if u'SMITHFIELD, STEVE' in names:
        indx = names.index(u'SMITHFIELD, STEVE')
        steve.append(xldata[indx][4])
    else:
        steve.append(0)

    if u'JOHNSON, JEFF' in names:
        indx = names.index(u'JOHNSON, JEFF')
        jeff.append(xldata[indx][4])
    else:
        jeff.append(0)
    wb.Close()

fpjeffsteve.write ("File,Jeff,Steve\n")
for i in range(len(xlxsfiles)):
    fpjeffsteve.write ("%s,%0.2f,%0.2f\n"%(xlxsfiles[i],jeff[i],steve[i]))
print ("Wrote jeffsteve.csv")
excel.Application.Quit()
