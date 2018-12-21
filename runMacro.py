import os
import win32com.client

#Launch Excel and Open Wrkbook
xl=win32com.client.Dispatch("Excel.Application")  
xl.Workbooks.Open(Filename="C:\Users\khanhdangnguyen\Documents\MyMacro.xlsm") #opens workbook in readonly mode. 

#Run Macro
xl.Application.Run("MyMacro.xlsm!Module1.ConvertToNewExcel") 

#Save Document and Quit.
xl.Application.Save()
xl.Application.Quit() 

#Cleanup the com reference. 
del xl