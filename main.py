import openpyxl,os,json
import datetime
import csv 
path = 'September Birthday list - test data.xlsx'
filePath = os.path.abspath(path)
CsvFileName = 'birthday.csv'
book = openpyxl.load_workbook(filePath)
user_data = book.get_sheet_by_name(str('Sheet3'))

# Get the four days preceding the current date. 
Current_Date = datetime.datetime.today()
Previous_Date = Current_Date - datetime.timedelta(days=4)
preDate = str(Previous_Date).split(" ")[0]

for x in range(1,user_data.max_row):
    BirthdayName = str(user_data[x][0].value)
    getBirDate = str(user_data[x][6].value)
 
    if getBirDate != "None" and getBirDate !="Birthday" and BirthdayName != "None" and BirthdayName != "Name":
        storeBirthdayDate = getBirDate.split(" ")[0]
        if preDate == storeBirthdayDate:
            # print(storeBirthdayDate,BirthdayName)
            insertIntoCsv = [storeBirthdayDate,BirthdayName]
      
            with open(CsvFileName,'a',newline='') as csvfile:
                writer = csv.writer(csvfile)
                # # writer.writerow(["Incomming Birthday","Name"])
                writer.writerow(insertIntoCsv)  
    else:
        print("False")



