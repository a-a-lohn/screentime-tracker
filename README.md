# screentime_tracker

*What the program does*
update.py is a simple program that can convert the output data of the app-usage Android app [StayFree](https://play.google.com/store/apps/details?id=com.burockgames.timeclocker&hl=en_CA) into an Excel Spreadsheet format. All historical app-usage data can be recorded and easily manipulated using an Excel Pivot Table or other tools.

*How to use it*
1. You must have the [StayFree](https://play.google.com/store/apps/details?id=com.burockgames.timeclocker&hl=en_CA) app, which can be downloaded for free from the Android Pay Store.
2. StayFree records your app-usage data on your phone and has the ability to export a CSV file containing your app usage for the past seven days at a time. To use this screentime tracking program, a copy of this CSV must be exported:
	1. Open StayFree on your phone and press **Reports** on the bottom bar
	2. Press the three dots in the top right corner of the screen, then press **Export to CSV**
	3. Press OK on the pop-up message
	4. If you have Excel set up on your phone, StayFree should create an Excel document titled **HistoryReport.xlsx**. You will likely see a message saying that the file is an older format and can only be saved to a copy. Press **Save a copy** and save it in this repo under ./data, overwriting the current HistoryReport.xlsx file.
	Note: In order to be able to see all your historical app-usage data, you must repeat this process at the end of the day every week. The program is set to throw an error if the CSV is updated on any day other than a Friday, but this can be changed.
3. Overwrite the file **./data/AppData.xlsx** with a blank Excel file of the same name. Populate the file's *first row only* according to how it is shown in this example. Rename the sheet to **Data**.
4. Make sure you are running Python3.x, and install the requirements (**pip -r install -r requirements.txt**).

After this initial set up make sure **AppData.xslx** is closed. Then run **update.py** to populate the spreadsheet with your app usage. Note that the **HistoryReport.xslx** file must be updated weekly, but **update.py** can run at any time.