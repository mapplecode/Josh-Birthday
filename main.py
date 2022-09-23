from genericpath import exists
import openpyxl,os
import datetime
import csv 
import schedule
import time

CsvFileName = 'birthday.csv'
if CsvFileName is not exists:
    with open(CsvFileName,'a',newline='') as f:
        writer = csv.writer(f)
    
def getFile():
    paths = "file"
    getFolderPath = os.path.abspath(paths)          
    for file in os.listdir(getFolderPath):
        if file.endswith(".xlsx"):
            file_path = f"{getFolderPath}\{file}"
            book = openpyxl.load_workbook(file_path)
            user_data = book.get_sheet_by_name(str('Sheet3'))
    return user_data

def main():
    user_data = getFile()
    os.remove(CsvFileName)
    # Get the four days preceding the current date. 
    Current_Date = datetime.datetime.today()
    Previous_Date = Current_Date + datetime.timedelta(days=4)
    preDate = str(Previous_Date).split(" ")[0]
    formateCurrentDate = str(Current_Date).split(" ")[0]
    with open(CsvFileName,'a',newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['DOB','Name'])
        for x in range(1,user_data.max_row):
            BirthdayName = str(user_data[x][0].value)
            getBirDate = str(user_data[x][6].value)
        
            if getBirDate != "None" and getBirDate !="Birthday" and BirthdayName != "None" and BirthdayName != "Name":
                storeBirthdayDate = getBirDate.split(" ")[0]
                if storeBirthdayDate>=preDate and storeBirthdayDate<formateCurrentDate:
                    insertIntoCsv = [storeBirthdayDate,BirthdayName]     
                    writer.writerow(insertIntoCsv) 
        return writer

if __name__ == "__main__":
    schedule.every(1).minutes.do(main)
    while True:
        schedule.run_pending()
        time.sleep(1)
   
