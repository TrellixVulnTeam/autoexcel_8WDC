# Example
'''
Modify here
'''
libraryExcelPath = r'C:\Users\khanhdangnguyen\Documents\ExcelPackage'
workingPath = r'C:\Users\khanhdangnguyen\Documents\checkDoc\DMS\MCU_Modeling\DEV'
versionColumn = 'B'
startRow = 3
sheetCheck = 'Version'
import numbers

def is_number(s):
	try:
		float(s)
		return True
	except ValueError:
		return False

def checkInFolder(folder):
	os.chdir('.\\'+folder)
	logFile.write(os.getcwd()+'\n')
	fileList = os.listdir(".")
	for item in fileList:
		print(item)
		if os.path.isdir('.\\'+item):
			checkInFolder(item)
			os.chdir('..')
			logFile.write(os.getcwd()+'\n')
		elif item.endswith('.xlsx'):
			excelItem = item.ljust(40)
			logFile.write(excelItem)
			wb = openpyxl.load_workbook(item)
			sheet = wb.get_sheet_by_name(sheetCheck)
			if (sheet == None):
				logFile.write('Err'.rjust(5)+'\n')
			else:
				i = startRow
				version = 'None'
				# while(sheet[versionColumn+str(i)].value != None):
					# version = str(sheet[versionColumn+str(i)].value)
					# i += 1
				for row in range(startRow,sheet.max_row + 1):
					if is_number(str(sheet[versionColumn+str(row)].value)):
						version = str(sheet[versionColumn+str(row)].value)
				logFile.write(version.rjust(5)+'\n')
		elif item.endswith('.xls'):
			print(str(item))
			excelItem = item.ljust(40)
			logFile.write(excelItem)			
			logFile.write('Old'.rjust(5)+'\n')
	# os.chdir('..')


import os
currentPath = os.getcwd()
print('Current working directory: '+os.getcwd())
print('Move to library directory: '+libraryExcelPath)
os.chdir(libraryExcelPath)
os.chdir('.\et_xmlfile-1.0.1')
import et_xmlfile
os.chdir('..\jdcal-1.4')
import jdcal
os.chdir('..\openpyxl')
import openpyxl
os.chdir(r'..')

os.chdir(workingPath)
print('Move to directory: '+workingPath)

fileList = os.listdir(".")
import pprint
#pprint.pprint(fileList)

logFile = open(r'.\log.txt','w')
logFile.write(os.getcwd()+'\n')
logFile.close()

logFile = open(r'.\log.txt','a')
checkInFolder('.')			
logFile.close()
	
os.chdir(workingPath)
print('Current working directory: '+workingPath)
logFile = open(r'.\log.txt')
print(logFile.read())
logFile.close()

print('Move back to library directory: '+currentPath)
os.chdir(currentPath)
