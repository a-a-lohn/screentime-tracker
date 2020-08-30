from openpyxl import load_workbook, workbook
from re import search, findall
from math import floor
import datetime
# do this to shorten method call
from datetime import date, timedelta

''' things to modify if adding a new column to beginning Data:
-if statement to determine if running initAppInfo()
-initAppInfo() (2 locations)
-for loop populating app_cells_in_data
-for loop iterating through times to calculate total for day
'''

hist_path = r"/mnt/c/Users/Aaron/OneDrive - McGill University/Programming/HistoryReport.xlsx"
data_path = r"/mnt/c/Users/Aaron/OneDrive - McGill University/Programming/screentime_tracker/AppData.xlsx"

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

def initTimeInfo():
    wsd['A1'] = 'time_info'
    wsd['$A$2'] = 'Season'
    wsd['$B$2'] = 'DOTW'
    wsd['$C$2'] = 'Date'

def disregardOlderSheets(most_recent_data_date, sheet_dates):
    # convert from datetime.datetime to datetime.date
    most_recent_data_date = date(most_recent_data_date.year, most_recent_data_date.month, most_recent_data_date.day)
    get_later_date = (date[0] for date in sheet_dates)
    i = 0
    while i < len(sheet_dates) and next(get_later_date) <= most_recent_data_date:
        i = i+1
    # delete sheets from sheet_dates that are from dates prior to the most recent date in Data, i.e.
    # their values have already been added
    if i > 0:
        del sheet_dates[:i]
    return sheet_dates

def initAppInfo(sheet_dates):
    wsd['F1'] = 'app_info'
    # get the oldest sheet
    wbh.active = wbh.get_sheet_by_name(sheet_dates[0][1])
    wsh = wbh.active
    # Skip over first cell, which is blank and the last few, which have other info
    get_cols = wsh.iter_cols(min_row=2, max_row=wsh.max_row-4, values_only=True)
    # dealing with column of apps
    app_names_in_sheet = next(get_cols)
    for app, i in zip(app_names_in_sheet, range(0, len(app_names_in_sheet))):
        wsd.cell(column=i+6, row=2, value=app)

# UNUSED
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
    # converts a string from a cell in HistoryReport into a Datetime.time object

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
    return datetime.time(h,m,s)
    
# UNUSED
def intToExcelCol(num):
    if num <= 26:
        return str(chr(num+64))
    first = chr(floor(num / 26) + 64)
    second = chr(num % 27 + 65)
    return str(first) + str(second)

# run this code to generate an updated graph of screentime usage
'''
TODO:
-add 'season' input option, asking if same season applies throughout and option to use past season
-allow data to be added from any day, not just Friday
-add groups so apps do not all show up individually
'''

def main():
    # get list of (date, name) tuples for each sheet from sheetDates()
    sheet_dates = []
    sheetDates(sheet_dates)
    # sort dates from the earliest to the most recent
    sheet_dates.sort(key=lambda tup: tup[0])

    # initialize time_info
    if wsd['A1'].value != 'time_info':
        initTimeInfo()
    time_info_cells = wsd['A2:C2']
    
    # get the most recent date recorded in Data. time_info_cells has all cell data in its first tuple
    most_recent_data_date = wsd.cell(column=time_info_cells[0][2].column, row=wsd.max_row).value
    
    # if Data sheet is not empty, use latest date in sheet to overwrite sheet_dates with only the newest sheets
    if most_recent_data_date != "Date":
        sheet_dates = disregardOlderSheets(most_recent_data_date, sheet_dates)
    #print(sheet_dates)
    
    # initialize app_info--should only be done if there are no app name column headers in Data, for whatever reason
    if wsd['F1'].value != 'app_info':
        initAppInfo(sheet_dates)
    
    # get a list with all the app cells in the Data sheet
    app_cells_in_data = []
    for i in range(6, wsd.max_column+1):
        app_cells_in_data.append(wsd.cell(column=i, row=2))
    
    # make a list for app values
    app_values_in_data = [app.value for app in app_cells_in_data]

    # first empty row that will be populated in Data sheet
    new_row_data = wsd.max_row + 1
    new_col_data = wsd.max_column + 1

    # hardcoded. Change manually when appropriate
    season = "Summer"

    #dict to convert DOTW number to abbreviation
    days = {0:"M", 1:"Tu", 2:"W", 3:"Th", 4:"F", 5:"Sa", 6:"Su"}

    # every time a new app is added, this will be set to true to increment new_col_data
    new_app = False

    for sheet in sheet_dates:
        #print(sheet)
        app_cells_to_append = []
        app_values_to_append = []

        # 1) start by adding the time info
        for cell in time_info_cells[0]:
            if cell.value == "Season":
                for i in range(0,7):
                    wsd.cell(column=cell.column, row=new_row_data+i, value=season)
            if cell.value == "DOTW":
                for i in range(0,7):
                    date_minus_i = sheet[0] - timedelta(days=6-i)
                    # the value is found by converting the required date (date_minus_i) to a date object,
                    # then using its weekday() attribute to query the days dictionary to return the desired
                    # string abbreviation for the day
                    wsd.cell(column=cell.column, row=new_row_data+i, value=days[date(date_minus_i.year, \
                        date_minus_i.month, date_minus_i.day).weekday()])
            if cell.value == "Date":
                for i in range(0,7):
                    wsd.cell(column=cell.column, row=new_row_data+i, value=sheet[0]-timedelta(days=6-i))

        # 2) now deal with app info. Start by working with the first hist sheet to add its data
        wbh.active = wbh.get_sheet_by_name(sheet[1])
        wsh = wbh.active

        # Skip over first cell, which is blank and the last few, which have other info
        get_cols = wsh.iter_cols(min_row=2, max_row=wsh.max_row-4, values_only=True)
        # dealing with column of apps
        app_names_in_sheet = next(get_cols)
        #print(app_names_in_sheet)

        # generate the app times, app by app, day by day
        app_row = wsh.iter_rows(min_row=2, max_row=wsh.max_row-4, min_col=2, max_col=8, values_only=True)

        for app in app_names_in_sheet:
            if new_app == True:
                new_col_data += 1
                new_app = False
            
            if app in app_values_in_data:
                #print(app)
                app_times = next(app_row)
                # get values for time spent, getting all 7 values for a given app per iteration
                for i, time in zip(range(0,7), app_times):
                    time = convertToTime(time)
                    #print(app, i, time)
                    #print(wsd.cell(column=app_cells_in_data[app_values_in_data.index(app)].column, \
                    #        row=new_row_data+i).value)
                    #print(app_cells_in_data[app_values_in_data.index(app)].column, new_row_data+i)

                    # if the top cell of the column to which we are about to write is empty, go ahead
                    if wsd.cell(column=app_cells_in_data[app_values_in_data.index(app)].column, \
                            row=new_row_data+i).value is None:
                        #print("will write")
                        # write in the time data under the column header corresponding to the app name,
                        # over 7 rows
                        wsd.cell(column=app_cells_in_data[app_values_in_data.index(app)].column, \
                            row=new_row_data+i, value=time)
                    
                    # the cell will not be empty if the app name is an already downloaded duplicate. In this case, find it and
                    # write data over there instead
                    elif app_values_in_data.count(app) > 1:
                        # insert data at next occurrence of app name; start searching through app list,
                        # starting one position after first occurrence
                        wsd.cell(column=app_cells_in_data[app_values_in_data.index(app, \
                            app_values_in_data.index(app)+1)].column, row=new_row_data+i, value=time)
                        
                    # if this is a new app with the same name as an existing app, add it as a new app column header with the times
                    else:
                        #print("new duplicate")
                        # write in the app name as a column header (but only once when i==0, not 7 times)
                        if i == 0:
                            wsd.cell(column=new_col_data, row=2, value=app)
                            new_app = True
                        
                        # write in its values
                        wsd.cell(column=new_col_data, row=new_row_data+i, value=time)

                        # append all new apps after iteratiing through them so that next sheet has updated list
                        app_cells_to_append.append(wsd.cell(column=new_col_data, row=2))
                        app_values_to_append.append(wsd.cell(column=new_col_data, row=2).value)

            # if the app was only downloaded this past week and is not a new duplicate
            else:
                # write in the app name as a column header
                wsd.cell(column=new_col_data, row=2, value=app)
                new_app = True

                # write in its values
                app_times = next(app_row)
                for i, time in zip(range(0,7), app_times):
                    time = convertToTime(time)
                    wsd.cell(column=new_col_data, row=new_row_data+i, value=time)
                
                app_cells_to_append.append(wsd.cell(column=new_col_data, row=2))
                app_values_to_append.append(wsd.cell(column=new_col_data, row=2).value)
        
        # 3) finally, deal with total and running avg columns for each day
        for i in range(0,7):
            total = timedelta()
            for time in wsd.iter_cols(min_row=new_row_data+i, max_row=new_row_data+i, \
                min_col=6, max_col=new_col_data, values_only=True):
                if time[0] is not None:
                    total += timedelta(hours=time[0].hour, minutes=time[0].minute, seconds=time[0].second)
            
            # take the last running avg value, multiply by num of days it was calculated for, add newest day total
            # and divide but new num of days it is being calculated for (i.e. 1 more day)

            # when first implemented, this if was needed since the last running avg value was of type datetime.time
            # instead of timedelta (last_run_avg assignment was inside else statement).
            # Since Aug 21 2020, ALL running avg and total values are of type timedelta
            #if isinstance(wsd.cell(column=5, row=new_row_data+i-1).value, datetime.time):
            #    print("true")
            #    last_run_avg = timedelta(hours=wsd.cell(column=5, row=new_row_data+i-1).value.hour, \
            #        minutes=wsd.cell(column=5, row=new_row_data+i-1).value.minute, \
            #        seconds=wsd.cell(column=5, row=new_row_data+i-1).value.second)
            #else:
            #    print("false")
            
            last_run_avg = wsd.cell(column=5, row=new_row_data+i-1).value
            mult_by = new_row_data+i-3
            div_by = new_row_data+i-2
            run_avg = (last_run_avg * mult_by + total) / div_by
            #print(run_avg)
            wsd.cell(column=4, row=new_row_data+i, value=total)
            wsd.cell(column=5, row=new_row_data+i, value=run_avg)

        app_cells_in_data = app_cells_in_data + app_cells_to_append
        app_values_in_data = app_values_in_data + app_values_to_append
        new_row_data = new_row_data + 7

        '''
        # lengthen app_info defined name range to include new apps from sheet, if new apps were added
        if new_app_counter + counter_at_start != new_app_counter:
            defined = wbd.defined_names['app_names']
            # dests is a tuple generator of (worksheet title, cell range)
            dests = defined.destinations
            _, coord = next(dests)
            first_cell = str(coord[0:4])
            last_cell = str(wsd.cell(column=wsd.max_column, row=2).coordinate)
            wbd.defined_names.delete('app_names')
            # replace app_names with larger range including new apps for next iteration
            new_range = workbook.defined_name.DefinedName('app_names', attr_text='Sheet!' + first_cell + ':' \
                        + last_cell)
            wbd.defined_names.append(new_range)'''
    
    # this is to deal with a bug that causes 00:00:00 times to appear as negative values in spreadsheet
    # Note that since the total and running avg cells are of type timedelta, the bug fix is NOT applied there, so
    # on days when phone is not used at all and total value is 00:00:00, the bug may occur and must be fixed manually
    for row in wsd.iter_rows(min_row=3, min_col=6, max_col=wsd.max_column, max_row=wsd.max_row):
        for cell in row:
            # convert from datetime.datetime to datetime.time
            if cell.value is not None:
                cell_time = datetime.time(cell.value.hour, cell.value.minute, cell.value.second)
                #print(cell_time)
                if cell_time == datetime.time(0,0,0):
                    cell.value = datetime.time(0,0,0)

    wbd.save(data_path)

if __name__ == "__main__":
    main()