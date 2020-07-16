# run this code to generate an updated graph of screentime usage

from openpyxl import load_workbook
from re import search, findall
from math import floor
from datetime import date, time, timedelta

hist_path = "screentime_tracker/HistoryReport.xlsx"
data_path = "screentime_tracker/AppData.xlsx"
# wb is short for workbook
wbh = load_workbook(hist_path)
wbd = load_workbook(data_path)
wbd.active = wbd.get_sheet_by_name("Data")
wsd = wbd.active

def sheetDates(sheet_dates):
    # creates a tupled list of all sheets generated from Samsung app in the form (date, name), where
    # date comes from their sheet names and name is the sheet name
    
    sheet_names = (name for name in wbh.sheetnames)
    for name in sheet_names:
        # having trouble using backreferences, just repeating the same group instead
        if search("^\d{2}-\d{2}-\d{4}_\d{2}-\d{2}-\d{2}$", name):
            # make sure all sheets have been generated on Fridays, if not throw an Exception
            if date(int(name[6:10]), int(name[3:5]), int(name[0:2])).weekday() != 4:
                raise Exception("Not all sheets were generated on a Friday")
            # append date object from sheet name with the sheet name. Date objects are in the form y,m,d
            sheet_dates.append((date(int(name[6:10]), int(name[3:5]), int(name[0:2])), name))

def defName(defined_name_in_wb):
    # returns a tuple containing all the cell values present in the Data sheet in a defined name

    defined = wbd.defined_names[defined_name_in_wb]
    # dests is a tuple generator of (worksheet title, cell range)
    dests = defined.destinations
    _, coord = next(dests)
    # Strangely enough, ws[cell_range_str] returns a single tuple with all the cells representing app names
    # inside the first inner tuple
    return wsd[coord][0]

def convertToTime(time_str):
    time_str = str(time_str)
    h,m,s = 0,0,0
    # if unit of time (h/m/s) is not present in cell (i.e. was not used long enough or was used for exactly
    # one minute, perhaps), mark as 0
    if search("\d{1,2}h", time_str):
        h = int(findall("\d{1,2}h", time_str)[0][:-1])
    if search("\d{1,2}m", time_str):
        m = int(findall("\d{1,2}m", time_str)[0][:-1])
    if search("\d{1,2}s", time_str):
        s = int(findall("\d{1,2}s", time_str)[0][:-1])
    return time(h,m,s)
    
# SHOULD NOT NEED
def intToExcelCol(num):
    if num <= 26:
        return str(chr(num+64))
    first = chr(floor(num / 26) + 64)
    second = chr(num % 27 + 65)
    return str(first) + str(second)

def main():
    # get list of (date, name) tuples for each sheet
    sheet_dates = []
    sheetDates(sheet_dates)
    # sort dates from the earliest to the most recent
    sheet_dates.sort(key=lambda tup: tup[1])

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
    
    # assuming order of time_info cells is Season, DOTW, Date but not assuming their coordinates in the sheet
    
    time_info_cells = defName('time_info')
    # get a tuple with all the app cells in the Data sheet
    app_names_in_data = defName('app_names')
    # make a list for app values
    app_values_in_data = [app.value for app in app_names_in_data]

    # first empty row that will be populated in Data sheet
    top_row = wsd.max_row + 1

    season = "Summer"
    # how can I take in input?
    #season = input("What season is it? If you do not enter one, the last season in the table will be used.")
    #if not season:
    #    season = 
    days = ["Sa", "Su", "M", "Tu", "W", "Th", "F"]
    for sheet in sheet_dates:
        # start by adding the time info
        for cell in time_info_cells:
            if cell.value == "Season":
                for i in range(0,7):
                    wsd.cell(column=cell.column, row=top_row+i, value=season)
            if cell.value == "DOTW":
                for i, day in zip(range(0,7), days):
                    wsd.cell(column=cell.column, row=top_row+i, value=day)
            if cell.value == "Date":
                for i in range(0,7):
                    wsd.cell(column=cell.column, row=top_row+i, value=sheet[0]-timedelta(days=6-i))

        # iterate over sheets from the least recent sheet the most
        wbh.active = wbh.get_sheet_by_name(sheet[1])
        wsh = wbh.active
        # max_row row value is a total row, so skip it

        #CHANGE THIS BACK TO VALUES ONLY
        get_cols = wsh.iter_cols(min_row=0, max_row=wsh.max_row-1, values_only=True)
        # dealing with column of apps. Skip over first cell, which is blank
        app_names_in_sheet = next(get_cols)[1:]

        # generate the app times, app by app
        app_times = wsh.iter_rows(min_row=2, max_row=wsh.max_row-1, min_col=2, max_col=8, values_only=True)
        for app in app_names_in_sheet:
            #*** deal with case of new apps later*** Will have to edit app_names defined name
            if app in app_values_in_data:
                # get values for time spent, getting all 7 values for a given app per iteration
                # for duplicate named case
                #if app_values_in_data.count(app.value) > 1:

                for i, time in zip(range(0,7), next(app_times)):
                    time = convertToTime(time)
                    # if the top cell of the column to which we are about to write is empty, go ahead
                    if wsd.cell(column=app_names_in_data[app_values_in_data.index(app)].column, \
                            row=top_row+i).value is None:
                        # write in the time data under the column header corresponding to the app name,
                        # over 7 rows
                        wsd.cell(column=app_names_in_data[app_values_in_data.index(app)].column, \
                            row=top_row+i, value=time)
                    # the cell will not be empty if the app name is a duplicate. In this case, find it and
                    # write data over there instead
                    else:
                        if app_values_in_data.count(app) > 1:
                            # insert data at next occurrence of app name; start searching through app list
                            # starting one position after first occurrence
                            wsd.cell(column=app_names_in_data[app_values_in_data.index(app, \
                                app_values_in_data.index(app)+1)].column, row=top_row+i, value=time)
                        # else: ***case of a new app***

    wbd.save(data_path)

    '''
    TODO:
    -add data from app-generated sheets to Data table column by column, adding new columns when necessary
    -format a Total column to capture total time from all app columns
    -delete sheets with added data

    #remove sheets from sheet_dates
    '''
if __name__ == "__main__":
    main()