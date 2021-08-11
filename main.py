from openpyxl import load_workbook
import json

READ_FILE = "DaytonInformationRajNew10AUG2020.xlsx"
WRITE_FILE = "Packaging_Matrix_UOM_Template_revE - DORMAN USE ONLY.xlsx"

workbook1 = load_workbook(READ_FILE, read_only=True, data_only=True)
sheet1 = workbook1["PKG Data New"]

data = []
for row in sheet1.iter_rows(min_row=2, min_col=1, max_col=7, values_only=True):
    if row[0] is None:
        break
    new_data = dict()
    new_data["partnum"] = row[0]
    new_data["length"] = row[4]
    new_data["width"] = row[5]
    new_data["height"] = row[6]
    data.append(new_data)

workbook2 = load_workbook(WRITE_FILE, data_only=True)
sheet2 = workbook2["Alt UoM Matrix"]

count = 0
for row in sheet2.iter_rows(min_row=14, min_col=3, max_col=9):
    if row[0].value == "H":
        row[1].value = data[count]["partnum"]
    elif row[2].value == "EA":
        row[4].value = data[count]["length"]
        row[5].value = data[count]["width"]
        row[6].value = data[count]["height"]
        count += 1

workbook2.save(filename=WRITE_FILE)