import win32com.client
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
workbook = excel.Workbooks.Open(r'E:\excel_workspace\autoexcel_repo\autoexcel\simple_example.xlsx')
sheet = workbook.Worksheets(1) #workbook.Sheets('Sheet1').Select(); 
sheet = xlApp.ActiveSheet
sheet.Cells(1,1).Value="Hello"
workbook.Save()
workbook.Close()
excel.Quit()
sheet = None
book = None
excel.Quit()
excel = None