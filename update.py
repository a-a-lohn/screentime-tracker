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
            #if date(int(name[6:10]), int(name[3:5]), int(name[0:2])).weekday() != 4:
            #    raise Exception("Not all sheets were generated on a Friday")
            # append date object from sheet name with the sheet name. Date objects are in the form y,m,d
            sheet_dates.append((date(int(name[6:10]), int(name[3:5]), int(name[0:2])), name))

def defName(defined_name_in_wb):
    # returns a tuple containing all the cell values present in the Data sheet in a defined name

    defined = wbd.defined_names[defined_name_in_wb]
    # dests is a tuple generator of (worksheet title, cell range)
    dests = defined.destinations
    _, coord = next(dests)
    # Strangely enough, wsd[coord] returns a single tuple with all the cells representing app names
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

    # get the most recent date recorded in Data
    # assumes order of time_info cells is Season, DOTW, Date but does not assume their coordinates in the sheet
    time_info_cells = defName('time_info')
    most_recent_data_date = wsd.cell(column=time_info_cells[2].column, row=wsd.max_row).value
    # only continue if Data sheet was originally populated with some date
    if most_recent_data_date != "Date":
        # convert from datetime.datetime to datetime.date
        most_recent_data_date = date(most_recent_data_date.year, most_recent_data_date.month, most_recent_data_date.day)
        get_later_date = (date[0] for date in sheet_dates)
        i = 0
        while i < len(sheet_dates) and next(get_later_date) <= most_recent_data_date:
            i = i+1
            print(i)
        # delete sheets from sheet_dates that are from dates prior to the most recent date in Data, i.e.
        # their values have already been added
        if i > 0:
            del sheet_dates[:i]
    
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
    days = {0:"M", 1:"Tu", 2:"W", 3:"Th", 4:"F", 5:"Sa", 6:"Su"}
    new_app_counter = 0
    for sheet in sheet_dates:
        counter_at_start = new_app_counter
        # start by adding the time info
        for cell in time_info_cells:
            if cell.value == "Season":
                for i in range(0,7):
                    wsd.cell(column=cell.column, row=top_row+i, value=season)
            if cell.value == "DOTW":
                for i in range(0,7):
                    date_minus_i = sheet[0] - timedelta(days=6-i)
                    # the value is found by converting the required date (date_minus_i) to a date object,
                    # then using its weekday() attribute to query the days dictionary to return the desired
                    # string abbreviation for the day
                    wsd.cell(column=cell.column, row=top_row+i, value=days[date(date_minus_i.year, \
                        date_minus_i.month, date_minus_i.day).weekday()])
            if cell.value == "Date":
                for i in range(0,7):
                    wsd.cell(column=cell.column, row=top_row+i, value=sheet[0]-timedelta(days=6-i))

        # iterate over sheets from the least recent sheet the most
        wbh.active = wbh.get_sheet_by_name(sheet[1])
        wsh = wbh.active
        # max_row row value is a total row, so skip it

        get_cols = wsh.iter_cols(min_row=0, max_row=wsh.max_row-1, values_only=True)
        # dealing with column of apps. Skip over first cell, which is blank
        app_names_in_sheet = next(get_cols)[1:]

        # generate the app times, app by app
        app_times = wsh.iter_rows(min_row=2, max_row=wsh.max_row-1, min_col=2, max_col=8, values_only=True)
        for app in app_names_in_sheet:
            #*** deal with case of new apps later*** Will have to edit app_names defined name
            if app in app_values_in_data:
                # get values for time spent, getting all 7 values for a given app per iteration
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
            # if the app was only downloaded this past week and was not in the data sheet previously
            else:
                # write in the app name as a column header
                wsd.cell(column=wsd.max_column+1, row=top_row, value=app)
                for i, time in zip(range(0,7), next(app_times)):
                    time = convertToTime(time)
                    # max_column is now column with new app name
                    wsd.cell(column=wsd.max_column, row=top_row+i, value=time)
                new_app_counter += 1
        # lengthen app_names defined name range to include new apps from sheet, if new apps were added
        if new_app_counter + counter_at_start != new_app_counter:
            defined = wbd.defined_names['app_names']
            # dests is a tuple generator of (worksheet title, cell range)
            dests = defined.destinations
            _, coord = next(dests)
            first_cell = str(coord[0:3])
            last_cell = str(wsd.cell(column=wbd.max_col, row=2).coordinate)
            wbd.defined_names.delete('app_names')
            # replace app_names with larger range including new apps for next iteration
            new_range = wbd.defined_name.DefinedName('app_names', attr_text='Sheet!' + first_cell + ':'/
                                 + last_cell)
            wb.defined_names.append(new_range)




    wbd.save(data_path)

    '''
    TODO:
    -add data from app-generated sheets to Data table column by column, adding new columns when necessary
    -format a Total column to capture total time from all app columns
    -add 'season' input option, asking if same season applies throughout and option to use past season
    '''
if __name__ == "__main__":
    main()