from openpyxl import load_workbook
import csv

csv_file_path = 'Data'
# Open the CSV file in read mode
fileName = "Data.csv"
fileName2 = "Data2.csv"
dict = {}
dict2 ={}

uniqueIdentifier = "CRN"
uniqueIdentifierCol = 0
with open(fileName, 'r', newline='') as csvfile:
    # Create a CSV reader object
    csv_reader = csv.reader(csvfile)




    # Iterate over each row in the CSV file
    for i, row in enumerate(csv_reader):
        if i == 0:  # This if statement finds the index of the given unique identifier in column 1
            for j, v in enumerate(row):
                if v == uniqueIdentifier:
                    uniqueIdentifierCol = j
        else:
            dict[row[uniqueIdentifierCol]] = row
# dicT[25474]= 202402,CPS,100,001,25474,,A,CLASS,D,GLSN,310,,,W,,,,,10:50 AM,12:05 PM,20,21,-1,0,0,0,1,04,Introduction to Cybersecurity,1,"Renna, James", ,  ,0,,,,
# dict[25795]= 202402,CPS,100,002,25795,,A,CLASS,D,GLSN,310,M,,,,,,,10:50 AM,12:05 PM,20,20,0,0,0,0,1,04,Introduction to Cybersecurity,1,"Renna, James", ,  ,0,,,,


#print(dict)

with open(fileName2, 'r', newline='') as csvfile2:
    # Create a CSV reader object
    csv_reader2 = csv.reader(csvfile2)
    for i, row in enumerate(csv_reader2):
        if i == 0:  # This if statement finds the index of the given unique identifier in the first column
            for j, v in enumerate(row):
                if v == uniqueIdentifier:
                    uniqueIdentifierCol = j
        else:
            dict2[row[uniqueIdentifierCol]] = row


for key in dict.keys():
    if key in dict2.keys():
        if dict[key] != dict2[key]:
            print(f"Differences found for key '{key}':")
            print(f"Dictionary 1 value: {dict[key]}")
            print(f"Dictionary 2 value: {dict2[key]}")
    else:
        print(f"Key '{key}' only exists in Dictionary 1")

for key in dict2.keys():
    if key not in dict.keys():
        print(f"Key '{key}' only exists in Dictionary 2")


