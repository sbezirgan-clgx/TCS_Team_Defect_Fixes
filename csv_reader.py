import csv
import os
import csv
import my_date

# csv file name
date_info = my_date.My_date()
filename = "out1.csv"



# initializing the titles and rows list
fields = []
rows = []


with open(filename, "r") as inf:
    # creating a csv reader object

    csvreader = csv.reader(inf)

    # extracting field names through first row
    fields = next(csvreader)

    # extracting each data row one by one
    for row in csvreader:
        rows.append(row)

    # get total number of rows
    a =(csvreader.line_num)
    print(a)

inf.close()
