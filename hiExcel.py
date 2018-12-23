#
# Open an existing workbook
#
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(r'E:\excel_workspace\autoexcel_repo\autoexcel\simple_example.xlsx')
# Alternately, specify the full path to the workbook 
# wb = excel.Workbooks.Open(r'C:\myfiles\excel\workbook2.xlsx')
ws = wb.Worksheets.Add()
ws.Name = "MyNewSheet"
wb.Save()
excel.Visible = True
excel.Application.Quit()