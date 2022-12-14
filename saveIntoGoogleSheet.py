import gspread,datetime,os,openpyxl,schedule,time
from google.oauth2.service_account import Credentials
scope = ['https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive']
creds = Credentials.from_service_account_file("creds.json", scopes=scope)
client = gspread.authorize(creds)
google_sh = client.open("Upcomming Birthday")
sheet1 = google_sh.worksheet('Sheet1')
def getFile():
    paths = "file"
    getFolderPath = os.path.abspath(paths)          
    for file in os.listdir(getFolderPath):
        if file.endswith(".xlsx"):
            file_path = f"{getFolderPath}\{file}"
            book = openpyxl.load_workbook(file_path)
            user_data = book.get_sheet_by_name(str('Sheet3'))
    return user_data
def updateGsheet(range,ExcelRecords):
    updateSh = sheet1.update(range,ExcelRecords)
    return updateSh
def main():
    user_data = getFile()
    sheetRange = 2
    # Get the four days preceding the current date. 
    Current_Date = datetime.datetime.today()
    Previous_Date = Current_Date + datetime.timedelta(days=4)
    preDate = str(Previous_Date).split(" ")[0]
    formateCurrentDate = str(Current_Date).split(" ")[0]
    for x in range(1,user_data.max_row):
        BirthdayName = str(user_data[x][0].value)
        getBirDate = str(user_data[x][6].value)
        if getBirDate != "None" and getBirDate !="Birthday" and BirthdayName != "None" and BirthdayName != "Name":
            storeBirthdayDate = getBirDate.split(" ")[0]
            if storeBirthdayDate<=preDate and storeBirthdayDate>formateCurrentDate:
                insertIntoGoogleSheet = [storeBirthdayDate,BirthdayName]
                ranges = f"A{sheetRange}:B{sheetRange}"
                updateGsheet(ranges,[insertIntoGoogleSheet])
                sheetRange+=1          
if __name__ == "__main__":
    schedule.every(1).minutes.do(main)
    while True:
        schedule.run_pending()
        time.sleep(1)
   




    