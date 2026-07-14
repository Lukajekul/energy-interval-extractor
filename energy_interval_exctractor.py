# IMPORTING ALL THE NECESARY LIBRARIES
from openpyxl import Workbook
from datetime import datetime, timedelta
from dateutil import tz
from openpyxl.workbook import Workbook
import sys

#CREATING THE WORKBOOK OBJECT
workbook = Workbook()


# CONTAINS THE MAIN LOGIC FOR THE EXTRACTING ALGORITHM
def read_file(path):
    # DECLARING ALL THE GLOBAL VARIABLES
    global sheet
    global fullList
    global warningDictionary
    global firstRun
    global measurementPointId
    global directionValue

    # INITIALIZES THE TIMEZONES AND INTERVAL
    from_zone = tz.gettz('UTC')
    to_zone = tz.tzlocal()
    interval = timedelta(minutes=15)

    # INITIALIZES THE VARIABLES
    firstRun = True
    fullList = []
    rawRows = []
    warningDictionary  = {}
    try:

        # OPENS THE FILE AND READS IT LINE BY LINE
        with open(path) as file:
            for line in file:
                splitLine = line.split(";")
                rawRows.append(splitLine)

            # INITIALIZES THE SHEET AND NAMES IT BASED ON THE LAST 3 NUMBERS IN MPID COLUMN
            sheet = workbook.create_sheet(title=f"MP_{rawRows[1][1][-3:]}")
            rawRows = rawRows[1:]
            for row in rawRows:

                # CONDITION CHECK FOR THE WARNING LABEL
                if row[5] != "L3":
                    if row[5] not in warningDictionary:
                        warningDictionary[row[5]] = 1
                    else:
                        warningDictionary[row[5]] += 1

                # CONVERTS THE EUROPEAN STANDARD FOR NUMBERS TO THE ONE THAT THE INTERPRETER USES
                kwhValue = float(row[3].replace(".", ""))
                kwhValue /= 1000000

                # SAVES THE MEASURMENT POINT ID FOR LATER
                row.pop(0)
                if firstRun:
                    measurementPointId = row[0]
                    firstRun = False
                row.pop(0)

                # TAKES DATE PARSES IT SO IT CAN ADD A 15 MINUTE INTERVAL TO IT
                dateTime = datetime.strptime(row[0], "%Y-%m-%d %H:%M:%S.%f")
                dateTime = dateTime.replace(tzinfo=from_zone)
                dateTime = dateTime.astimezone(to_zone)

                # ADD THE INTERVAL AND TURNS BOTH NUMBERS INTO TIME OBJECTS
                toTime = dateTime + interval
                fromTime = datetime.strftime(dateTime, "%Y-%m-%d %H:%M")
                toTime = datetime.strftime(toTime, "%Y-%m-%d %H:%M")

                # CHEKS WHAT DIRECTION THE ELECTRICITY HAS AND REMEMBERS IT
                if row[2] == 'r':
                    directionValue = "Recieved"
                elif row[2] == "s":
                    directionValue = "Sent"

                # SENDS ALL THE AQUIRED DATA INTO AN OTHER FUCNTION FOR ROW CREATINGON
                collect_row(fromTime, toTime, row[3], kwhValue, rawRows)

    # ERROR HANDELING
    except FileNotFoundError:
        print("One or more of the inputed paths are invalid.")
        sys.exit("Terminating program due to file error.")

# FILLS THE fullList WITH ALL THE ROWS AND WAIT FOR THE FILE TO BE READ FULLY
def collect_row(fromTime, toTime, method, kwhValue, rawRows):

    # CREATES A ROW LIST AND PASES IT INTO ANOTHER LIST
    row = [fromTime, toTime, method, kwhValue]
    fullList.append(row)

    # WHEN THE LIST IS SAME LEGHT AS THE ROW IN THE FILE IT SORTS THE IT BY from TIME AND ADDS HARDCOADED HEARED VALUES
    if len(fullList) == len(rawRows):
        sortedList = sorted(fullList, key=lambda x: x[0])
        sheet['A1'] = "From"
        sheet['B1'] = "To"
        sheet['C1'] = "MeasurmentHethod"
        sheet['D1'] = "Value in KWH"
        sheet['F1'] = "MeasumentPoint:"
        sheet['G1'] = measurementPointId
        write(sortedList)

# WRITES ALL THE LINES INTO THE FILE AND ADD THE WARNING LABEL
def write(sortedList):
    # GENERIC BEGINING OF EVERY WARNIGN
    warningMessage = "Warning! There have been"
    lenBegining = len(warningMessage)

    # FILS THE WARNING MESSAGE WITH WARNINGS
    for i in warningDictionary:
        warningMessage += f" {warningDictionary[i]} instances of {i},"
    for row in sortedList:
        sheet.append(row)

    # IF THE WARNING MESSAGE CAHNGED IN LENGHT IT ADDS IT TO THE .xlsx FILE
    if len(warningMessage) > lenBegining:
        sheet['F4'] = f"{warningMessage[:-1]}"

    # ADD THE LAST TWO HARDCOADED VALUES
    sheet['F2'] = "Direction:"
    sheet['G2'] = directionValue


# MAIN FUCTION THAT GETS THE NECESARY PATHS TO THE FILES CALLS THE read_file FUCNITON AND SAVES THE FILE AT THE END
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
    workbook.save(f"{fileName}.xlsx")

if __name__ == '__main__':
    main()