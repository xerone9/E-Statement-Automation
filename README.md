# E-Statement-Automation
This Automation Process is designed for a specific institute (It wont work for everyone)

Daily Work Folder should be presend on the desktop.

All PDF files (E-statements Askari Bank) present on the root of this folder will be combined and a combine PDF and Text File will get opened.
Copy all the contents in the PDF (ctrl + a) and paste it in the text file and save (You must save eitherwise the output report will be blanked). Then go to python application opened terminal and press enter. Note it will terminate All opened text files and Adobe reader.

Application will then read the text file and remove all the lines in the statement other then lines starting with "Clearing Outward" and "Cash Deposit" and will create 2 csv files.

1st file is for Indus University EMS that contains only 2 columns "Serial number" and "Receipt No" (without headers) the file will be saved by the name of date present on the statement e.g: 12Aug22.csv. Login to your LMS account and go to accounts, upload fee voucher and upload that csv file (after reviewing it) thier select date refresh and post.

2nd File is for reconciliation and will contains columns like "Serail number", "Branch Code", "Narration/Particulars", "Amount" (without headers)the file will be saved with the name of For Reconciliation + date e.g: For Reconciliation 12-Aug-22.csv

Both files will get saved in the Daily Work folder mentioned above

Then it will create a folder by the date of E-Statement and move all pdfs in that folder. and will delete combine output pdf file

System Requirements

Windows 7 sp1 or above

Appliation Requirements

1- Will only work with Daily Work Folder Present On desktop

2- That folder must contain a text file with name "New Text Document.txt" 
