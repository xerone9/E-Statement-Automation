from PyPDF2 import PdfFileMerger
import glob, os, re
import csv
import shutil
from pygrok import Grok


def listToStringWithoutBrackets(value):
    return str(value).replace('[', '').replace(']', '').replace(',', '').replace("'", "").replace('  ',
                                                                                                  '*').replace(' ',
                                                                                                               '')
def listToStringWithoutBrackets1(value):
    return str(value).replace('*', ' ').replace('\\n', '').replace('@', ',')


def listToStringWithoutBrackets2(value):
    return str(value).replace('[', '').replace(']', '').replace(',', '').replace("'", "")


def adjustmentToString(value):
    return str(value).replace(',', '@').replace('\\n', '')

def adjustmentToCSVDate(value):
    return str(value).replace('-', '')


def listRemoveDateAdditional(value):
    return str(value).replace('{','').replace('}','').replace("'","").replace(",","-").replace('day: ','').replace(" month: ","").replace(" year: ","")


############################ Will search for All PDFs in the folder and combine them ######################################

desktop = os.path.expanduser("~\desktop\\Daily Work\\")
os.chdir(desktop)
file_dict = {}
for file in glob.glob("*.pdf"):
    filepath = file
    if filepath.endswith((".pdf", ".PDF")):
            file_dict[file] = filepath
merger = PdfFileMerger(strict=False)

for k, v in file_dict.items():
    print(k, v)
    merger.append(v)

merger.write(desktop + "output.pdf")
merger.close()

################### Open the combine PDF and Notepad file (Copy full PDF and paste it on text file and save #######################

output = os.path.expanduser("~\desktop\\Daily Work\\output.pdf")
Text = os.path.expanduser("~\desktop\\Daily Work\\New Text Document.txt")
f = open(Text, "w")
f.write("Open output.pdf File and Paste data here. MAKE SURE TO SAVE DATA EITHER WISE THE FILE WILL BE BLANK!!!")
f.close()
os.startfile(Text)
os.startfile(output)
print("")
input("Make sure you had saved Text File. (IT WILL BE DONE MANUALLY CTRL + S)....")
os.system("TASKKILL /F /IM AcroRd32.exe")
os.system("TASKKILL /F /IM Notepad.exe")

########################## Will read the text file and remove all irrelevant lines just leave clearing outward and cash deposit lines ##########################

desktop = os.path.expanduser("~\desktop\\Daily Work\\New Text Document.txt")
date = filepath[-15] + filepath[-14] + filepath[-13] + filepath[-12] + filepath[-11] + filepath[-10] + filepath[-9] + filepath[-6] + filepath[-5]
date1 = filepath[-15] + filepath[-14] + filepath[-13] + filepath[-12] + filepath[-11] + filepath[-10] + filepath[-9] + filepath[-8] + filepath[-7] + filepath[-6] + filepath[-5]

with open(desktop, 'rt', encoding="ISO-8859-1") as f:
    data = f.readlines()
open(desktop, 'w').close()
for line in data:
    lineMan = str(line)
    dated = ""
    if line.__contains__('CASH DEPOSIT') or line.__contains__('CLEARING OUTWARD'):
        if len(line) > 11:
            dated = lineMan[0] + lineMan[1] + lineMan[2] + lineMan[3] + lineMan[4] + lineMan[5] + lineMan[6] + lineMan[
                7] + lineMan[8] + lineMan[9] + lineMan[10]
        else:
            dated = ""
        g = open(desktop, "a")
        if dated == date1:
            symbo = str(line)
            wordsec = adjustmentToString(symbo)
            word = list(wordsec)
            word.pop(0)
            word.pop(0)
            word.pop(0)
            word.pop(0)
            word.pop(0)
            word.pop(0)
            word.pop(0)
            word.pop(0)
            word.pop(0)
            word.pop(0)
            word.pop(0)
            word.pop(0)
            words = str(word)
            word = listToStringWithoutBrackets(words)
            line = listToStringWithoutBrackets1(word)
            g.write(line + "\n")
        else:
            g.write(line)
        g.close()


fo = open(output, "wb")
#Close opend file
fo.close()
os.remove(output)
# os.startfile(Text)

########################## Creating CSV file with same date name to upload on Indus EMS (Contains Serial and Receipt No Only ######################

csvDate = adjustmentToCSVDate(date)
f = open(Text, "r+")
c = open(csvDate + '.csv', 'w', newline='')
writer = csv.writer(c)
serial = 0
for x in f:
    receipts = [int(s) for s in x.split() if s.isdigit()]
    serial += 1

    if x.startswith("CASH DEPOSIT"):
        line.strip()
        receipts.pop(0)
        receipts.insert(0, serial)
        writer.writerow(receipts)
    else:
        line.strip()
        receipts.insert(0, serial)
        writer.writerow(receipts)
c.close()
f.close()

################## Creating Exel File For Reconciliation Cotaining all columns (date, narration, amount etc) ##########################
i = open(Text, "r")
j = open('For Reconciliation ' + date1 + '.csv', 'w', newline='')
writer = csv.writer(j)
serial = 0
dateFix = []
for x in i:
    serial += 1
    if x.startswith("CASH DEPOSIT"):
        id = x.strip()
        input_string = id
        date_pattern = '%{MONTHDAY:day}-%{MONTH:month}-%{YEAR:year}'
        grok = Grok(date_pattern)
        dateSelect =listRemoveDateAdditional(grok.match(input_string)) + ' '
        if dateSelect not in dateFix:
            dateFix.append(dateSelect)
        try:
            result = id.split(dateSelect)[1]
            convertingCashDeposit = id.split(dateSelect)[0]
            withBranch = [convertingCashDeposit[0:16]]
            branch = str(withBranch).replace('CASH DEPOSIT', 'Branch ')
            removingCashDeposit = id.split(dateSelect)[0]
            removedCashDeposit = removingCashDeposit[16:]
            receipts = [removedCashDeposit]
            finalResult = result.split(' ')[0]
            # receipts.pop(0)
            receipts.insert(0, listToStringWithoutBrackets2(branch))
            receipts.insert(0, serial)
            receipts.insert(4, dateSelect)
            receipts.insert(5, finalResult)
            writer.writerow(receipts)
        except IndexError:
            dateCorrect = str(dateFix[0])
            dateSelect = dateCorrect[-9:-1] + ' '
            result = id.split(dateSelect)[1]
            convertingCashDeposit = id.split(dateSelect)[0]
            withBranch = [convertingCashDeposit[0:16]]
            branch = str(withBranch).replace('CASH DEPOSIT', 'Branch ')
            removingCashDeposit = id.split(dateSelect)[0]
            removedCashDeposit = removingCashDeposit[16:]
            receipts = [removedCashDeposit]
            finalResult = result.split(' ')[0]
            # receipts.pop(0)
            receipts.insert(0, listToStringWithoutBrackets2(branch))
            receipts.insert(0, serial)
            receipts.insert(4, dateCorrect)
            receipts.insert(5, finalResult)
            writer.writerow(receipts)
    else:
        id = x.strip()
        input_string = id
        date_pattern = '%{MONTHDAY:day}-%{MONTH:month}-%{YEAR:year}'
        grok = Grok(date_pattern)
        dateSelect = listRemoveDateAdditional(grok.match(input_string)) + ' '
        result = id.split(dateSelect)[1]
        clearingOutwardFilter = id[0:16]
        clearingOutward = [clearingOutwardFilter]
        extractingClearingOutwardCode = id.split(dateSelect)[0]
        extractedCodeFilter = id[17:]
        extractedCode = extractedCodeFilter.split(dateSelect)[0]
        receipts = [extractedCode]
        finalResult = result.split(' ')[0]
        receipts.insert(0, listToStringWithoutBrackets2(clearingOutward))
        receipts.insert(0, serial)
        receipts.insert(4, dateSelect)
        receipts.insert(5, finalResult)
        writer.writerow(receipts)
j.close()
f.close()

########################### Creating Relevant Date Folder and moving same dates PDF in it ######################

makeFolder = os.path.expanduser("~\desktop\\Daily Work\\" + date)
os.mkdir(makeFolder)
desktop = os.path.expanduser("~\desktop\\Daily Work\\")




for file in glob.glob("*.pdf"):
    filepath = file
    shutil.move(desktop + "\\" + file,
             makeFolder + "\\" + file)

os.startfile(csvDate + '.csv')