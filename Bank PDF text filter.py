from PyPDF2 import PdfMerger
from PyPDF2 import PdfReader
from PyPDF2 import PdfWriter
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


    ############################ Renaming Account Statements ######################################

accounts = []

with open("accounts.ini") as file:
    data = file.readlines()
    for line in data:
        values = str(line.strip())
        accounts.append(values)

passwords = []

with open("passwords.ini") as file:
    data = file.readlines()
    for line in data:
        values = str(line.strip())
        passwords.append(values)

desktop = os.path.expanduser("~\desktop\\Daily Work\\")

############################ Decrypting Files if are Password Protected ######################################

for filename in os.listdir(desktop):
    if filename.endswith((".pdf", ".PDF")):
        encrypted_file = desktop + filename
        decrypted_file = desktop + "unlocked - " + filename
        # reader = PdfReader(desktop + filename)
        with open(encrypted_file, 'rb') as pdf_file:
            pdf_reader = PdfReader(pdf_file)
            if pdf_reader.is_encrypted:
                # Loop through each password and try to decrypt the file
                for password in passwords:
                    if pdf_reader.decrypt(password) == 1:
                        # If the password is correct, create a new PDF file without password
                        pdf_writer = PdfWriter()
                        for page_num in range(len(pdf_reader.pages)):
                            pdf_writer.add_page(pdf_reader.pages[page_num])
                        with open(decrypted_file, 'wb') as output_file:
                            pdf_writer.write(output_file)
                        pdf_reader.stream.close()
                        os.remove(desktop + filename)
                        break
                else:
                    print("No Password Found For: " + encrypted_file)
            else:
                pass

for filename in os.listdir(desktop):
    if filename.endswith((".pdf", ".PDF")):
        with open(desktop + filename, 'rb') as pdf_file:
            pdf_reader = PdfReader(pdf_file)
            if pdf_reader.is_encrypted:
                pdf_reader.stream.close()
                os.remove(desktop + filename)


for filename in os.listdir(desktop):
    if filename.endswith((".pdf", ".PDF")):
        reader = PdfReader(desktop + filename)
        pages = len(reader.pages)
        account_number = ""
        statement_date = ""
        for i in range(pages):
            page = reader.pages[i]
            data = str(page.extract_text())
            if data.__contains__("Pak Rupees"):
                account_number = data.split("Pak Rupees")[1].split(" Account #")[0]
                statement_date = data.split(" ** Closing Balance **")[0][-11:]
                if statement_date.__contains__("DB"):
                    statement_date = data.split(" ** Opening Balance ** ")[1].split("Value Date")[0]
                os.rename(str(desktop + filename), desktop + account_number + " - " + statement_date + ".pdf")
                break

for filename in os.listdir(desktop):
    if filename.endswith((".pdf", ".PDF")):
        account = str(filename).split(" - ")[0]
        if account not in accounts:
            os.remove(desktop + filename)


    ############################ Will search for All PDFs in the folder and combine them ######################################

file_dict = {}
os.chdir(desktop)
for file in glob.glob("*.pdf"):
    filepath = file
    if filepath.endswith((".pdf", ".PDF")):
            file_dict[file] = filepath
merger = PdfMerger(strict=False)

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


########################### Moving CSV Files To Relevant Folder ######################

for file in glob.glob("*.csv"):
    filepath = file
    shutil.move(desktop + "\\" + file,
             "csv_vault\\" + file)

os.startfile("csv_vault\\" + csvDate + '.csv')