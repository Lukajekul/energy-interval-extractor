from openpyxl import Workbook, load_workbook
from datetime import datetime, timedelta
from dateutil import tz
from openpyxl.workbook import Workbook

workbook = Workbook()


def read_file(path):
    global page
    global fullList
    global warningDistionary
    global firstRun
    global MeasurementPoint
    global directionValue
    firstRun = True
    fullList = []
    openingList = []
    warningDistionary  = {}
    with open(path) as file:
        for line in file:
            splitLine = line.split(";")
            openingList.append(splitLine)
        page = workbook.create_sheet(title=f"MP_{openingList[1][1][-3:]}")
        openingList = openingList[1:]
        for row in openingList:
            if row[5] != "L3":
                if row[5] not in warningDistionary:
                    warningDistionary[row[5]] = 1
                else: 
                    warningDistionary[row[5]] += 1          
            
            valueNumberKWH = float(row[3].replace(".", ""))
            valueNumberKWH /= 1000000

            row.pop(0)
            if firstRun:
                MeasurementPoint = row[0]
                firstRun = False
            row.pop(0)
            

            dateTime = datetime.strptime(row[0], "%Y-%m-%d %H:%M:%S.%f")
            from_zone = tz.gettz('UTC')
            to_zone = tz.tzlocal()
            dateTime = dateTime.replace(tzinfo=from_zone)
            dateTime = dateTime.astimezone(to_zone)


            durationObject = timedelta(minutes=15)
            
            
            toTime = dateTime + durationObject

            fromTime = datetime.strftime(dateTime, "%Y-%m-%d %H:%M")
            toTime = datetime.strftime(toTime, "%Y-%m-%d %H:%M")

            
            if row[2] == 'r':
                directionValue = "Recieved"
            elif row[2] == "s":
                directionValue = "Sent"

            fillInputList(fromTime, toTime, row[3], valueNumberKWH, openingList)

def fillInputList(fromTime, toTime, measurment, valueNumberKWH, openingList):
    row = [fromTime, toTime, measurment, valueNumberKWH]
    fullList.append(row)

    if len(fullList) == len(openingList):
        sortedList = sorted(fullList, key=lambda x: x[1])
        page['A1'] = "From"
        page['B1'] = "To"
        page['C1'] = "MeasurmentHethod"
        page['D1'] = "Value in KWH"
        page['F1'] = "MeasumentPoint:"
        page['G1'] = MeasurementPoint
        write(sortedList)

def write(sortedList):
    warningMessage = "Warning! There have been"
    lenBegining = len(warningMessage)
    for i in warningDistionary:
        warningMessage += f" {warningDistionary[i]} instances of {i},"
    for row in sortedList:
        page.append(row)
    if len(warningMessage) > lenBegining:
        page['F4'] = f"{warningMessage[:-1]}"
    page['F2'] = "Direction:"
    page['G2'] = directionValue

def main ():
    listOfPaths = []
    numberOfFiles = int(input("How many files are you working with? -->"))
    if numberOfFiles == 1:
        path = str(input("What is the path to the file? -->"))
    else: 
        for i in range(numberOfFiles):
            path = str(input("What is the path to the file? -->"))
            listOfPaths.append(path)
    if numberOfFiles == 1:
        read_file(path)
    else:
        for i in listOfPaths:
            read_file(i)
    fileName = str(input("What would you like to name your file? -->"))
    del workbook["Sheet"]
    workbook.save(f"../{fileName}.xlsx")

if __name__ == '__main__':
    main()