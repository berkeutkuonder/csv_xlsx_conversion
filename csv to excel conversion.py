# -*- coding: utf-8 -*-
"""
Created on Tue Oct 31 21:23:17 2017

@author: BerkeUtkuOnder
"""
import os
import csv
import xlsxwriter

data = []

def main_loop():
    while True:
        csvname = input("Enter the name of the file: ")
        if csvname[-4:] != ".csv":
            print("The ending should be '.csv'!")
            continue
        break
    load_csv_list(csvname)
    while True:
        excelname = input("Enter a name for the new file: ")
        if excelname[-5:] != ".xlsx":
            print("The ending should be '.xlsx'!")
            continue
        break  
    save_excel_list(excelname)
    print("Conversion complete!")

def load_csv_list(csvname):
    if os.access(csvname,os.F_OK):
        f = open(csvname)
        for row in csv.reader(f):
            data.append(row)
        f.close()

def save_excel_list(excelname):
    workbook = xlsxwriter.Workbook(excelname,{'strings_to_numbers': True})
    worksheet = workbook.add_worksheet("Sheet1")
    row = 0
    col = 0
    for a in data:
        b = len(a)
        for num in range(b):
            worksheet.write(row, col, a[num])
            col += 1
        row += 1
        col = 0

if __name__ == '__main__':
    main_loop()