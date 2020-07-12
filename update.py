# run this code to generate an updated graph of screentime usage

#import os

#from flask import Flask
#from models import db, Book
import xlrd
import re
from datetime import date

def main():
    path = "screentime_tracker/HistoryReport.xlsx"
    # wb is short for workbook
    wb = xlrd.open_workbook(path)
    sheet_names = (name for name in wb.sheet_names())
    sheet_dates = []
    for name in sheet_names:
        # having trouble using backreferences
        if re.search("^\d{2}-\d{2}-\d{4}_\d{2}-\d{2}-\d{2}$", name):
            # append date object from sheet name. Date objects are in the form y,m,d
            sheet_dates.append(date(int(name[6:10]), int(name[3:5]), int(name[0:2])))
    # sort dates from the most recent one to the least
    sheet_dates.sort(reverse=True)
    data_sheet = wb.sheet_by_name("Data")
    last_row = data_sheet.row_values(data_sheet.nrows-1, start_colx=1, end_colx=4)
    # indices correspond to labelling of d, m, y in table on Data sheet
    most_recent_data_date = date(int(last_row[2]), int(last_row[1]), int(last_row[0]))
    get_later_date = (date for date in sheet_dates)
    i = -1
    while get_later_date.next() > most_recent_data_date:
        i = i+1
    
if __name__ == "__main__":
    main()