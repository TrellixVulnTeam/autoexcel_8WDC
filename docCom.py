import win32com.client as win32
import os
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False

# for infile in glob.glob( os.path.join('', '*.docx') ):
# doc = word.Documents.Open(os.getcwd()+'\'+infile)
wdFormatUnicodeText = 7
doc = word.Documents.Open(r'D:\python_workspace\Automate_the_Boring_Stuff_onlinematerials_v.2\INT-MCS-18018_DTSG_U2.docx')

# word.ActiveDocument.SaveAs('doc.txt',FileFormat=win32com.client.constants.wdFormatText)
doc.SaveAs('D:\python_workspace\Automate_the_Boring_Stuff_onlinematerials_v.2\doc.txt', wdFormatUnicodeText)
if not doc.CheckGrammar:
	print "Did not pass the grammar and spelling check"
else:
	print "Pass spelling check"
	
# Number of features used (at least 3 to get the full grade)
tableFeatures = 0
 
print('There are '+str(doc.Tables.Count)+' tables')
# print(doc.Content.Text)
doc.Close()
# doc.Save
word.Application.Quit(-1)