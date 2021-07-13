from openpyxl import load_workbook

wb = load_workbook("xl.xlsx", data_only=True)

#load_ws = load_wb['Sheet1']

ws = wb[wb.sheetnames[0]]
commonrail = []

print("A")
all_values = []
for row in ws.rows:
    row_value = []
    for cell in row:
        row_value.append(cell.value)
    all_values.append(row_value)
print(all_values)


#L1-Primary infos
print(all_values[1])

#L2-Default infos
print(all_values[3][:5])

#L3-LIST
l3_values = []
for row in ws.iter_rows(min_row=6):
    if row[0].value == None : break
    row_value = []
    for cell in row[0:3]:
        row_value.append(cell.value)
    l3_values.append(row_value)
print(l3_values)
