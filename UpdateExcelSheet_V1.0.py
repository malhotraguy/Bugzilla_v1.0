import openpyxl
from openpyxl.styles import Font
from ExcelFunc import BugzillaExcelPop
wb=openpyxl.Workbook()
#wb.save('BugReports_updated.xlsx')

sheet=wb.active
colum=["Bug ID","Product","Component","Creation_time","Last_change_time","Duplicates"]
for colNum in range(1,7):
    sheet.cell(row=1, column=colNum).font = Font(size=11,name='Calibri', bold=True)
    sheet.cell(row=1,column=colNum).value=colum[colNum-1]

wb.save('BugReports_updated.xlsx')
BugzillaExcelPop("BugReports.xlsx",'BugReports_updated.xlsx',colum[1:6])
