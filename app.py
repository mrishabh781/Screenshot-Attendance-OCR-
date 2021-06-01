import smtplib
from random import randint
import gspread
from gspread.utils import rowcol_to_a1
import time
import cv2
import datetime
from oauth2client.service_account import ServiceAccountCredentials
import glob
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import cv2
import numpy as np
import pandas as pd
import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell

# variables
scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

date = str(datetime.date.today())
print(date)
#date = '2021-01-06'
final_attendance = [date]


def update_sheet(final_attendance):
    creds = ServiceAccountCredentials.from_json_keyfile_name("cred.json", scope)
    client = gspread.authorize(creds)
    sheet = client.open("CompilerDesign_2020_Attendance").get_worksheet(0) # seting up the values to fetch the sheet
    data = sheet.get_all_records()
    length = len(data[0])
    last_roll = len(data)
    h = rowcol_to_a1(1, length + 1) + ':' + rowcol_to_a1(last_roll + 1, length + 1)
    print(h)
    cell_list = sheet.range(h)
    # print(cell_list)
    for cell, atend in zip(cell_list, final_attendance):
        cell.value = atend
        sheet.update_cells(cell_list)
    print("updated")


def get_names(fileNames):
    names = set({})
    for i in fileNames:
        image=cv2.imread(i)
        width = image.shape[1]
        cropped_image = image[:, int(width* 0.865):-int(width/32)]
        #cv2.imwrite(f'cro{i}.png',cropped_image)
        #print(image.shape[1])
        text = pytesseract.image_to_string(cropped_image, lang='eng')
        text = text.lower().strip().replace('.','').split('\n')
        names.update(text)
    return names

def get_namesw(fileNames):
    names = set({})
    for i in fileNames:
        image=cv2.imread(i)
        width = image.shape[1]
        cropped_image = image[:, int(width* 0.832):-int(width/22)]
        #cv2.imwrite(f'croped_{i}.png',cropped_image)
        #print(image.shape[1])
        text = pytesseract.image_to_string(cropped_image, lang='eng')
        text = text.lower().strip().replace('.','').split('\n')
        names.update(text)
    return names



names = glob.glob('*.png')
name = get_names(names)

excel = pd.read_excel('excel.xlsx')
#excel.drop_column(['Unnamed: 0'])

#print(dict(excel))
column = list(excel.columns)

stu_name = list(excel[column[1]])

present = []
not_reco = []
#date = input('input date')

for i in stu_name:
    for j in name:
        if i.lower().strip() in j:
            present.append(i)
            break
print(sorted(name))
print(sorted(present))
print(len(present))


print("not recognised : ",sorted(not_reco))

final_attend = []

for i in stu_name:
    if i in present:
        final_attend.append(1)
    else:
        final_attend.append(0)

#print(final_attend)

excel[date] = final_attend
writer = pd.ExcelWriter('excel.xlsx',engine='xlsxwriter')
excel.to_excel(writer,index=False)
writer.save()

final_attendance.extend(final_attend)

#update_sheet(final_attendance)
