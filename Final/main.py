def updateGsheet(range,ExcelRecords):
    updateSh = sheet1.update(range,ExcelRecords)
    return updateSh

def main():
    sheetRange = 2 
    # Get the four days preceding the current date. 
    Current_Date = datetime.datetime.today()
    Previous_Date = Current_Date + datetime.timedelta(days=4)
    preDate = str(Previous_Date).split(" ")[0]
    print(preDate,"====")
    formateCurrentDate = str(Current_Date).split(" ")[0]
    print(formateCurrentDate,"====")
    for getRecords in Sheet2Data:
     
        getBirDate = getRecords.get('Date of Birth (mm/dd)')
        BirthdayName = getRecords.get('Name')
        getBirDateReplace = getBirDate.replace("/","-")
        convertDate = datetime.datetime.strptime(getBirDateReplace,'%m-%d-%Y').strftime('%Y-%m-%d')
        if convertDate<=preDate and convertDate>formateCurrentDate:
            insertIntoGoogleSheet = [convertDate,BirthdayName]
            ranges = f"A{sheetRange}:B{sheetRange}"
            updateGsheet(ranges,[insertIntoGoogleSheet])
            sheetRange+=1
            
if __name__ == "__main__":
    import gspread,datetime
    from google.oauth2.service_account import Credentials
    scope = ['https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive']
    creds = Credentials.from_service_account_file("creds.json", scopes=scope)
    # creds = Credentials.from_service_account_file('/home/Stand4Socks/code/creds.json', scopes=scope)
    client = gspread.authorize(creds)
    google_sh = client.open("Birthday Automation Sheet")
    sheet1 = google_sh.worksheet('Sheet1')
    sheet2 = google_sh.worksheet('Sheet2')
    Sheet2Data = sheet2.get_all_records()
    data = sheet1.get_all_records()
    sheet1.delete_rows(2,20)
    main()



