from docx import Document
import csv

dates=[]
nature=[]
from_destination=[]
to_destination=[]
miles=[]

#add code for adding total miles

with open("miles.csv", "r") as file:
    reader = csv.reader(file)
    for i in range(8):
        next(reader)
    for row in reader: 
        #columns = [row[i] for i in columns_to_print]
        if row[0]=="" and row[1]=="":
            break
        dates.append(row[0])
        if len(row)>0 and row[1]!="":
            nature.append(row[1])
        if len(row)>2 and row[3]!="":
            from_destination.append(row[3])
        if len(row)>3 and row[4]!="":
            to_destination.append(row[4])
        if len(row)>4 and row[5]!="":
            miles.append(row[5])

doc = Document("Mileage_Template.docx")

table = doc.tables[0]

for row_index in range(6, len(nature)+6):
    row = table.rows[row_index]
    curr_row_index = row_index-6

    for i in range(len(row.cells)):
        if i==0 and len(dates)>curr_row_index and dates[curr_row_index]!="":
            row.cells[i].text=dates[curr_row_index]
        if i==2 and len(nature)>curr_row_index:
            row.cells[i].text=nature[curr_row_index]
        if i==5 and len(from_destination)>curr_row_index:
            row.cells[i].text=from_destination[curr_row_index]
        if i==7 and len(to_destination)>curr_row_index:
            row.cells[i].text=to_destination[curr_row_index]
        if i==11 and len(miles)>curr_row_index:
            row.cells[i].text=miles[curr_row_index]

doc.save("Completed_Mileage.docx")
