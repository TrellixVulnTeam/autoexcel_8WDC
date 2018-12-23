
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
# wb = excel.Workbooks.Add()
wb = excel.Workbooks.Open(r'E:\excel_workspace\autoexcel_repo\autoexcel\check_range.xlsx')
ws = wb.Worksheets("Sheet1")
used = ws.UsedRange
nrows = used.Row + used.Rows.Count - 1
ncols = used.Column + used.Columns.Count - 1
print(str(nrows))
print(str(ncols))
print(str(used.Row))
print(str(used.Column))
print(str(used.Rows.Count))
print(str(used.Columns.Count))
excel.Visible = False
excel.Application.Quit()