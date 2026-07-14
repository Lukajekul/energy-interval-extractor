from openpyxl import Workbook, load_workbook
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import csv
from dateutil import tz
import xlsxwriter
from openpyxl.workbook import Workbook

header = ["From", "To", "Measurment", "Value in KWH", "", "MeasurmentPoint:"]


FIRST_RUN = True
FIRST_PRINT_RUN = True

FULL_LIST = []

# workbookWrite = Workbook()
# sheetWrite = workbookWrite.active


workbook = Workbook()
page = workbook.active
#page.title = "Test"

def read_file(path):
    openingList = []
    with open(path) as file:
        for line in file:
            splitLine = line.split(";")
            openingList.append(splitLine)
        openingList = openingList[1:]
        for row in openingList:


            
            valueNumberKWH = float(row[3].replace(".", ""))
            valueNumberKWH /= 1000000

            row.pop(0)
            global MeasurementPoint
            global FIRST_RUN
            if FIRST_RUN:
                MeasurementPoint = row[0]
                FIRST_RUN = False
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

            global directionValue
            if row[2] == 'r':
                directionValue = "Recieved"
            elif row[2] == "s":
                directionValue = "Sent"

            fillInputList(fromTime, toTime, row[3], valueNumberKWH, openingList)

def fillInputList(fromTime, toTime, measurment, valueNumberKWH, openingList):
    global FIRST_PRINT_RUN
    global FULL_LIST
    if FIRST_PRINT_RUN:
        row = ["From", "To", "MeasurmentHethod", "Value in KWH", "", "MeasumentPoint:", MeasurementPoint]
        FULL_LIST.append(row)
        row = [fromTime, toTime, measurment, valueNumberKWH, "", "Direction:", directionValue]
        FULL_LIST.append(row)
        FIRST_PRINT_RUN = False
    else:
        row = [fromTime, toTime, measurment, valueNumberKWH]
        FULL_LIST.append(row)
    if len(FULL_LIST) == len(openingList):
        write(FULL_LIST)

def write(FULL_LIST):
    for row in FULL_LIST:
        page.append(row)
    workbook.save("../test.xlsx")

def main ():
    path = "../readings_quarter_hourly_02072026 071238.csv"
    read_file(path)

if __name__ == '__main__':
    main()