# run this code to generate an updated graph of screentime usage

from openpyxl import load_workbook
import re
from datetime import date

def main():
    path = "screentime_tracker/HistoryReport.xlsx"
    # wb is short for workbook
    wb = load_workbook(path)

    # create a tupled list of all sheets generated from Samsung app in the form (date, name), where
    # date comes from their sheet names and name is the sheet name
    sheet_names = (name for name in wb.sheetnames)
    sheet_dates = []
    for name in sheet_names:
        # having trouble using backreferences, just repeating the same group instead
        if re.search("^\d{2}-\d{2}-\d{4}_\d{2}-\d{2}-\d{2}$", name):
            # append date object from sheet name with the sheet name. Date objects are in the form y,m,d
            sheet_dates.append((date(int(name[6:10]), int(name[3:5]), int(name[0:2])), name))
    # sort dates from the most recent one to the least
    sheet_dates.sort(key=lambda tup: tup[1], reverse=True)

    # go to the Data sheet
    wb.active = wb.get_sheet_by_name("Data")
    ws = wb.active

    # get the most recent date recorded in Data

    # the following string represents the range of cells representing the most recent date in the Data sheet
    cell_range_str = "B" + str(ws.max_row) + ":" + "D" + str(ws.max_row)
    # convert to a date. Strangely enough, ws[cell_range_str] returns a single tuple with a triple tuple
    # inside that has the 3 cell values corresponding to the date
    y, m, d = ws[cell_range_str][0]
    most_recent_data_date = date(y.internal_value, m.internal_value, d.internal_value)


    get_later_date = (date[0] for date in sheet_dates)
    i = -1
    while next(get_later_date) > most_recent_data_date:
        i = i+1
    
if __name__ == "__main__":
    main()