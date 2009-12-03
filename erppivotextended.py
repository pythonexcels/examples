#
# erppivotextended.py:
# Load raw EPR data, clean up header info,
# insert additional data fields and build 5 pivot tables
#
import win32com.client as win32
win32c = win32.constants
import sys
import itertools
tablecount = itertools.count(1)

def addpivot(wb,sourcedata,title,filters=(),columns=(),
             rows=(),sumvalue=(),sortfield=""):
    """Build a pivot table using the provided source location data
    and specified fields
    """
    newsheet = wb.Sheets.Add()
    newsheet.Cells(1,1).Value = title
    newsheet.Cells(1,1).Font.Size = 16

    # Build the Pivot Table
    tname = "PivotTable%d"%tablecount.next()

    pc = wb.PivotCaches().Add(SourceType=win32c.xlDatabase,
                                 SourceData=sourcedata)
    pt = pc.CreatePivotTable(TableDestination="%s!R4C1"%newsheet.Name,
                             TableName=tname,
                             DefaultVersion=win32c.xlPivotTableVersion10)
    wb.Sheets(newsheet.Name).Select()
    wb.Sheets(newsheet.Name).Cells(3,1).Select()
    for fieldlist,fieldc in ((filters,win32c.xlPageField),
                            (columns,win32c.xlColumnField),
                            (rows,win32c.xlRowField)):
        for i,val in enumerate(fieldlist):
            wb.ActiveSheet.PivotTables(tname).PivotFields(val).Orientation = fieldc
            wb.ActiveSheet.PivotTables(tname).PivotFields(val).Position = i+1

    wb.ActiveSheet.PivotTables(tname).AddDataField(
        wb.ActiveSheet.PivotTables(tname).PivotFields(sumvalue[7:]),
        sumvalue,
        win32c.xlSum)
    if len(sortfield) != 0:
        wb.ActiveSheet.PivotTables(tname).PivotFields(sortfield[0]).AutoSort(sortfield[1], sumvalue)
    newsheet.Name = title

    # Uncomment the next command to limit output file size, but make sure
    # to click Refresh Data on the PivotTable toolbar to update the table
    # newsheet.PivotTables(tname).SaveData = False

    return tname

def runexcel():
    """Open the spreadsheet ABCDCatering.xls, clean it up,
    and add pivot tables
    """
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
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

    logolookup = {'Applied Materials':'AMAT', 'Electronic Arts':'EA',
                  'Hewlett-Packard':'HP', 'KLA-Tencor':'KLA'}
    if ("Company Name" in newdata[0]):
        cindx = newdata[0].index("Company Name")
        newdata[0][cindx+1:cindx+1] = ["Logo Name"]
        for rcnt in range(1,len(newdata)):
            if newdata[rcnt][cindx] in logolookup:
                newdata[rcnt][cindx+1:cindx+1] = [logolookup[newdata[rcnt][cindx]]]
            else:
                newname = newdata[rcnt][cindx].split()[0]
                newdata[rcnt][cindx+1:cindx+1] = [newname]
                logolookup[newdata[rcnt][cindx]] = newname
            
    foodlookup = {'Caesar Salad':'Salad', 'Cheese Pizza':'Pizza',
                  'Cheeseburger':'Burger', 'Chocolate Sundae':'Dessert',
                  'Churro':'Snack', 'Hamburger':'Burger', 'Hot Dog':'HotDog',
                  'Pepperoni Pizza':'Pizza', 'Potato Chips':'Snack',
                  'Soda':'Drink'}
    if ("Food Name" in newdata[0]):
        cindx = newdata[0].index("Food Name")
        newdata[0][cindx+1:cindx+1] = ["Food Category"]
        for rcnt in range(1,len(newdata)):
            if newdata[rcnt][cindx] in foodlookup:
                newdata[rcnt][cindx+1:cindx+1] = [foodlookup[newdata[rcnt][cindx]]]
            else:
                newdata[rcnt][cindx+1:cindx+1] = ['UNDEFINED']
            
    rowcnt = len(newdata)
    colcnt = len(newdata[0])
    wsnew = wb.Sheets.Add()
    wsnew.Range(wsnew.Cells(1,1),wsnew.Cells(rowcnt,colcnt)).Value = newdata
    wsnew.Columns.AutoFit()

    src = "%s!R1C1:R%dC%d"%(wsnew.Name,rowcnt,colcnt)

    # What were the total sales in each of the last four quarters?
    addpivot(wb,src,
             title="Sales by Quarter",
             filters=(),
             columns=(),
             rows=("Fiscal Quarter",),
             sumvalue="Sum of Net Booking",
             sortfield=())

    # What are the sales for each food item in each quarter?
    addpivot(wb,src,
             title="Sales by Food Item",
             filters=(),
             columns=("Food Name",),
             rows=("Fiscal Quarter",),
             sumvalue="Sum of Net Booking",
             sortfield=())

    # Who were the top 10 customers for ABCD Catering in 2009?
    addpivot(wb,src,
             title="Top 10 Customers",
             filters=(),
             columns=(),
             rows=("Company Name",),
             sumvalue="Sum of Net Booking",
             sortfield=("Company Name",win32c.xlDescending))

    # Who was the highest producing sales rep for the year?
    addpivot(wb,src,
             title="Top Sales Reps",
             filters=(),
             columns=(),
             rows=("Sales Rep Name","Company Name"),
             sumvalue="Sum of Net Booking",
             sortfield=("Sales Rep Name",win32c.xlDescending))

    # What food item had the highest unit sales in Q4?
    ptname = addpivot(wb,src,
             title="Unit Sales by Food",
             filters=("Fiscal Quarter",),
             columns=(),
             rows=("Food Name",),
             sumvalue="Sum of Quantity",
             sortfield=("Food Name",win32c.xlDescending))
    wb.Sheets("Unit Sales by Food").PivotTables(ptname).PivotFields("Fiscal Quarter").CurrentPage = "2009-Q4"

    # What food category had the highest unit sales in Q4?
    ptname = addpivot(wb,src,
             title="Unit Sales by Food Category",
             filters=("Fiscal Quarter",),
             columns=(),
             rows=("Food Category",),
             sumvalue="Sum of Quantity",
             sortfield=("Food Category",win32c.xlDescending))
    wb.Sheets("Unit Sales by Food Category").PivotTables(ptname).PivotFields("Fiscal Quarter").CurrentPage = "2009-Q4"

    if int(float(excel.Version)) >= 12:
        wb.SaveAs('newABCDCatering.xlsx',win32c.xlOpenXMLWorkbook)
    else:
        wb.SaveAs('newABCDCatering.xls')
    excel.Application.Quit()

if __name__ == "__main__":
    runexcel()
