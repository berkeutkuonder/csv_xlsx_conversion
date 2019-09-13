# -*- coding: utf-8 -*-
"""
Created on Mon Oct 16 21:00:37 2017

@author: BerkeUtkuOnder
"""
import openpyxl
import csv

data = []

def main_loop():
    while True:
        excelname = input("Enter a name for the new file: ")
        if excelname[-5:] != ".xlsx":
            print("The ending should be '.xlsx'!")
            continue
        break 
    print("You also have to enter the name of the sheet you want to convert")
    sheetname = input("Enter the name of the sheet: ")
    load_excel_list(excelname, sheetname)
    while True:
        csvname = input("Enter the name of the file: ")
        if csvname[-4:] != ".csv":
            print("The ending should be '.csv'!")
            continue
        break
    save_csv_list(csvname)
    print("Conversion complete!")

def load_excel_list(excelname, sheetname):
    wb = openpyxl.load_workbook(excelname)
    sheet = wb.get_sheet_by_name(sheetname)
    lit = []
    for row in sheet:
        for obj in row:
            lit.append(obj.value)
        data.append(lit)
        lit = []
    
def save_csv_list(csvname):
    f = open(csvname, "w", newline="")
    for item in data:
        csv.writer(f).writerow(item)
    f.close()

if __name__ == '__main__':
    main_loop()
    
#%%
