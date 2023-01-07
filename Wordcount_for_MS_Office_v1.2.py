# Importing libraries of MS Word, PowerPoint and Excel
import xlwings as xlw
import pandas as pd
import docx
import pptx
import win32com.client #To read.doc files
# Importing library of the operations with the Operating System and Platform
import platform as pf
import os
# Importing time library to get the ceation time of the files
import time
# Importing datetime to format the date of creation of the files
from datetime import datetime

# Asking the user to enter the path of the files to be processed
files_path = input("Enter path of the files:")

# Creating the workbook to hold the processed data
Excel_file_name = input("How do you want to name the Excel file of the processed data?")
'''
wb = xlw.Book()
wb.save()
pathxl = r'%s\%s.xlsx' %(files_path,Excel_file_name)
wb.save(pathxl)
sheet = wb.sheets['Sheet1']
'''

# Listing the files inside the directory and its subdirectories
res_files = {}
for (dir_path, dir_names, file_names) in os.walk(files_path):
    for fl in file_names:
        if fl.startswith("~"): # To exclude listing files used for temporary storage for an already existing file
            continue
        res_files[fl] = dir_path

'''
# Entering the files' names in the workbook from the list (res_files)
y = 1
for x in res_files:
    sheet["A"+str(y)].value = x
    y = y + 1
'''

# Function that finds the date of creation or last modification of the file
def creation_date(files_pth):
    x = datetime.fromtimestamp((os.path.getctime(files_pth)))
    y = x.strftime("%d/%m/%Y")
    return y

# Listing the dates of the files
dates = []
for x in res_files:
    pathfl = r'%s\%s' %(res_files[x],x)
    dates.append(creation_date(pathfl))

'''
# Entering the files' dates in the workbook from the list (dates)
y = 1
for x in dates:
    sheet["B"+str(y)].value = x
    y = y + 1
'''

# Funtion that finds the wordcount if the file is a text file
def word_count_txt(files_pth):
    # print(files_pth)
    word = win32com.client.Dispatch("Word.Application")
    word.visible = False
    wordcount = []
    for x in files_pth:
        # print("Good to here")
        y = list(files_pth).index(x)        
        if x.endswith(".docx") or x.endswith(".doc") or x.endswith(".txt"):
            wb = word.Documents.Open(x)
            content = wb.Range().Text        
            content_list = content.split(" ")
            if len(content_list) == 1 and len(content_list[0]) <= 1:
                wordcount.append(0)
                wb.Close()
            else:
                wordcount.append(len(content_list))
                wb.Close()
        else:
                wordcount.append(None)
    word.Quit()
    return wordcount

content_wordcount = []
list_paths = []
for x in res_files:
    #pathfl = r'%s\%s' %(res_files[x],x)   // Amended below to send list of paths, so that no need to call method many times, will call it just oncr
    list_paths.append(r'%s\%s' %(res_files[x],x))
# print("Reached here")
content_wordcount = word_count_txt(list_paths)

'''
# Entering the files' wordcounts in the workbook from the list (content_wordcount)
for y in res_files:
    z = list(res_files).index(y)
    sheet["C"+str(z + 1)].value = content_wordcount[z]
'''

# Open a workbook
wb = xlw.Book()
wb.save()
pathxl = r'%s\%s.xlsx' %(files_path,Excel_file_name)
wb.save(pathxl)
sheet = wb.sheets['Sheet1']

#  Entering all data from a dictionary
All_Data = {"File Name": res_files, "Date": dates, "Word count": content_wordcount}
pd.DataFrame(All_Data).to_excel(wb)
# Saving and closing the workbook
wb.save()
wb.close()