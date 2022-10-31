from genericpath import exists
import openpyxl,os
import datetime
import csv 
import schedule
# from datetime import datetime
import dateutil.parser

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
        writer.writerow(['Name','Home Address1','Home Address2','Home Address City','Home Address State','Home Address Postal Code','Home Address Country/Territory','Date of Birth (mm/dd)','Socks'])
        for x in range(1,user_data.max_row):
            BirthdayName = str(user_data[x][0].value)
            homeAddress1 = str(user_data[x][1].value)
            homeAddress2 = str(user_data[x][2].value)
            HomeAddressCity = str(user_data[x][3].value)
            HomeAddressState = str(user_data[x][4].value)
            HomeAddressPostalCode = str(user_data[x][5].value)
            HomeAddressCountryTerritory = str(user_data[x][6].value)
            getBirDate = str(user_data[x][7].value)
            Socks = str(user_data[x][8].value)

            if getBirDate != "None" and getBirDate !="Date of Birth (mm/dd)" and BirthdayName != "None" and BirthdayName != "Name" and homeAddress1 != "None" and homeAddress1 != "Home Address1" and homeAddress2 != "Home Address2"  and HomeAddressCity != "Home Address City"  and HomeAddressState != "Home Address State" and HomeAddressPostalCode != "Home Address Postal Code" and HomeAddressCountryTerritory != "Home Address Country/Territory" and Socks != "None" and Socks != "Socks":
                storeBirthdayDate = getBirDate.split(" ")[0]
                getBirDateReplace = storeBirthdayDate.replace("/","-")
                data = dateutil.parser.parse(getBirDateReplace)
                convertDate = str(data).split(" ")[0]
                if convertDate<=preDate and convertDate>formateCurrentDate:
                    insertIntoCsv = [BirthdayName,homeAddress1,homeAddress2,HomeAddressCity,HomeAddressState,HomeAddressPostalCode,HomeAddressCountryTerritory,storeBirthdayDate,Socks]     
                    writer.writerow(insertIntoCsv) 
        return writer

if __name__ == "__main__":
    main()

