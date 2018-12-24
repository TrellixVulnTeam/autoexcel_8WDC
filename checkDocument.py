import os
import win32com.client as win32
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
wdFormatUnicodeText = 7
versionRegex = re.compile(r'v\d.(\d)+')
currentPath = os.getcwd()
workingPath = r'C:\Users\khanhdangnguyen\Documents\03_Document'
libraryExcelPath = r'C:\Users\khanhdangnguyen\Documents\ExcelPackage'
print('Current working directory: '+os.getcwd())
print('Move to library directory: '+libraryExcelPath)
os.chdir(workingPath)
print('Move to directory: '+workingPath)

fileList = os.listdir(".")
outputFile = open(r'.\log.txt','w')

'''
Body start
'''

for file in fileList:
	if file.endswith('.docx'):
		print(file)
		outputFile.write(file.ljust(35))
		doc = word.Documents.Open(workingPath+'\\'+file)
		doc.SaveAs(workingPath+'\\'+'temp.txt', wdFormatUnicodeText)
		doc.Close()
		tempFile = open(workingPath+'\\'+'temp.txt')
		result = versionRegex.search('')
		while result is None:
			line = tempFile.readline()
			if line is None:
				break
			else:
				result = versionRegex.search(line)
		tempFile.close()
		if result is not None:
			outputFile.write(result.group().rjust(5))
		else:
			outputFile.write('Err'.rjust(5))
		outputFile.write('\n')
os.remove(workingPath+'\\'+'temp.txt')
del doc
del tempFile
word.Application.Quit(-1)
'''
Body end
'''
outputFile.close()
del outputFile
print('Move back to library directory: '+currentPath)
os.chdir(currentPath)