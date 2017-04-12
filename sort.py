#Script written to sort through PGP output file in CSV format
__author__ = 'Kamil Wilk'

import csv, operator
from datetime import datetime, timedelta

#CSVfile Header
csvheader = ["name", "primary_email_address", "desktop_lastseen", "last_access", 
            "mac_address", "hostname", "machine_id", "device_id", "ip_address", 
            "pgp_desktop_version","server_last_seen", "server_status_as_of", 
            "display_name", "case", "status"]

duplicates =    [""]


#Technical Contacts
techs = [""]


#hosts to ignore
irrelevantHosts = [""]


def main():
    #We grab our CSV file name
    csvToSort = raw_input("Enter the name of your input CSV file: ")
    basicSortName = csvToSort + "-Sorted"
    timeSortName = csvToSort + "-TimeSorted"
    duplicatesName = csvToSort + "-DuplicateGUIDs"

    #We add .csv if not included
    if ".csv" not in csvToSort:
        csvToSort += ".csv"
    if ".csv" not in basicSortName:
        basicSortName += ".csv"
    if ".csv" not in timeSortName:
        timeSortName += ".csv"
    if ".csv" not in duplicatesName:
        duplicatesName += ".csv"

    #Enter date of excel file
    excelDate = raw_input("Enter the .csv file date (mm/dd/yy): ")
    excelDate = datetime.strptime(excelDate, "%m/%d/%y")


    #Time Sorting parameter
    daysToSortBy = input("How recent(number of days) should results be? ")
    then = excelDate - timedelta(days = daysToSortBy+1)

    print "Sorting: " + csvToSort
    #Sorting the relevant columns
    print "Sorting with relevant columns..."
    with open(csvToSort ,"rb") as infile, open(basicSortName, "wb") as outfile, open (duplicatesName, "wb") as dupFile:
        rdr = csv.reader(infile)
        wtr = csv.writer(outfile)
        dupWtr = csv.writer(dupFile)
        dupWtr.writerow(csvheader)

        for r in rdr:
            try:
                #We check for duplicate GUIDs and add them to seperate list
                if (r[7].upper() in duplicates or r[12].upper() in duplicates):
                    dupWtr.writerow( (r[0], r[1], r[2], r[3], r[4], r[6].upper(), r[7], r[12], r[8], r[10], r[11], r[21], r[15], r[18], r[19]) )
                #If it wasn't a duplicate GUID, we add all C drives to list for sorting
                elif (r[15][0] == 'C' and r[6].upper() not in irrelevantHosts):
                    #if r[6].upper() not in irrelevantHosts:
                    wtr.writerow( (r[0], r[1], r[2], r[3], r[4], r[6].upper(), r[7], r[12], r[8], r[10], r[11], r[21], r[15], r[18], r[19]) )
            #We except IndexError so that we don't have a confusing if statements
            except IndexError:
                pass

    #Sort by hostname and dates
    print "Sorting by hostname and dates..."
    with open(basicSortName, "rb") as infile:
        rdr = csv.reader(infile)

        #this is an ugly solution to a weird problem being caused by the csv returning times as hour:minute:seconds-04/05
        sortedList = sorted(rdr, key = lambda row: datetime.strptime(row[10][0:19], "%Y-%m-%d %H:%M:%S"), reverse = True)
        sortedList = sorted(sortedList, key = lambda row: datetime.strptime(row[11][0:19], "%Y-%m-%d %H:%M:%S"), reverse = True)
        sortedList = sorted(sortedList, key = lambda row: datetime.strptime(row[2][0:19], "%Y-%m-%d %H:%M:%S"), reverse = True)
        sortedList = sorted(sortedList, key = lambda row: datetime.strptime(row[3][0:19], "%Y-%m-%d %H:%M:%S"), reverse = True)

        sortedList = sorted(sortedList, key=operator.itemgetter(5))


    with open(basicSortName, "wb") as outfile:
        wtr = csv.writer(outfile)

        for row in sortedList:
            wtr.writerow(row)


    #Sort with most recent non-technical contact user for each unique host
    print "Sorting most recent user for each unique host for final result..."
    with open(basicSortName, "rb") as infile:
        rdr = csv.reader(infile)

        rowlist = []
        for row in rdr:
            rowlist.append(row)


    with open(basicSortName, "wb") as outfile:
        wtr = csv.writer(outfile)
        wtr.writerow(csvheader)

        recentrow = []
        for r in rowlist:
            if (not recentrow):
                recentrow = r

            #if the current host is different then the recent one, write what was stored in the recent
            if (r[5] != recentrow [5]):
                wtr.writerow(recentrow)
                #update recent host
                recentrow = r

            #if the current host is the same as the recent one, see if we should replace the recent one
            elif (r[5] == recentrow[5]):
                #check for tech contacts
                if ( (r[0] not in techs) and (recentrow[0] in techs) ):
                    recentrow = r
                recentrow = isMoreRecent(r, recentrow, recentrow)


    #Time Sorted List
    with open(basicSortName, "rb") as infile, open(timeSortName, "wb") as outfile:
        rdr = csv.reader(infile)
        wtr = csv.writer(outfile)
        next(rdr, None)  #skip the header

        wtr.writerow(csvheader) #write a new header
        for row in rdr:
            cr6 = datetime.strptime(row[10][0:19], "%Y-%m-%d %H:%M:%S")
            cr7 = datetime.strptime(row[11][0:19], "%Y-%m-%d %H:%M:%S")
            #cr1 = datetime.strptime(row[2][0:19], "%Y-%m-%d %H:%M:%S")
            #cr2 = datetime.strptime(row[3][0:19], "%Y-%m-%d %H:%M:%S")
            if (cr6 >= then and cr7 >= then):
                wtr.writerow(row)

    print "Sorting Complete! "
    print "Basic Sorted Output file is: " + basicSortName
    print "Time Sorted Output file is: " + timeSortName
    print "Duplicate GUIDs Output file is: " + duplicatesName

#Function to check for the most recent date
def isMoreRecent(currentrow, previousrow, recentrow):
    cr6 = datetime.strptime(currentrow[10][0:19], "%Y-%m-%d %H:%M:%S")
    cr7 = datetime.strptime(currentrow[11][0:19], "%Y-%m-%d %H:%M:%S")
    cr1 = datetime.strptime(currentrow[2][0:19], "%Y-%m-%d %H:%M:%S")
    cr2 = datetime.strptime(currentrow[3][0:19], "%Y-%m-%d %H:%M:%S")
    pr6 = datetime.strptime(previousrow[10][0:19], "%Y-%m-%d %H:%M:%S")
    pr7 = datetime.strptime(previousrow[11][0:19], "%Y-%m-%d %H:%M:%S")
    pr1 = datetime.strptime(previousrow[2][0:19], "%Y-%m-%d %H:%M:%S")
    pr2 = datetime.strptime(previousrow[3][0:19], "%Y-%m-%d %H:%M:%S")

    if cr6 > pr6:
        recent = currentrow
    elif cr7 > pr7:
        recent = currentrow
    elif cr1 > pr1:
        recent = currentrow
    elif cr2 > pr2:
        recent = currentrow
    else:
        recent = recentrow
    return recent

if __name__ == '__main__':
    main()
