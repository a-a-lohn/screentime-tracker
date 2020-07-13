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
            # make sure all sheets have been generated on Fridays, if not throw an Exception
            if date(int(name[6:10]), int(name[3:5]), int(name[0:2])).weekday() != 4:
                raise Exception("Not all sheets were generated on a Friday")
            # append date object from sheet name with the sheet name. Date objects are in the form y,m,d
            sheet_dates.append((date(int(name[6:10]), int(name[3:5]), int(name[0:2])), name))
    # sort dates from the most recent one to the least
    sheet_dates.sort(key=lambda tup: tup[1], reverse=True)

    # go to the Data sheet
    wb.active = wb.get_sheet_by_name("Data")
    ws = wb.active
    '''
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
    # delete sheets from sheet_dates that are from dates prior to the most recent date in Data, i.e.
    # their values have already been added
    if i >= 0:
        del sheet_dates[i:]
    '''
    # MOST OF THE ABOVE CODE CAN BE REMOVED IF I JUST ADD A FEATURE AT THE END OF THIS CODE TO REMOVE SHEETS
    # WITH DATA THAT HAS ALREADY BEEN ADDED

    '''TODO:
    -add data from app-generated sheets to Data table column by column, adding new columns when necessary
    -format a Total column to capture total time from all app columns
    -delete sheets with added data


    #remove sheets from sheet_dates
    '''
if __name__ == "__main__":
    main()