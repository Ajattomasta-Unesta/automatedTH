import math
import datetime
from openpyxl import load_workbook

def load_data(filepath) :
    wb = load_workbook("xl.xlsx", data_only=True) #filepath
    #load_ws = load_wb['Sheet1']

    ws = wb[wb.sheetnames[0]]

    print("A")
    all_values = []
    for row in ws.rows:
        row_value = []
        for cell in row:
            row_value.append(cell.value)
        all_values.append(row_value)
    print(all_values)

    #L1-Primary infos
    L1 = all_values[1]
    print(L1)

    #L2-Default infos
    L2 = all_values[3][:5]
    print(L2)

    #L3-LIST
    L3 = []
    for row in ws.iter_rows(min_row=6):
        if row[0].value == None : break
        row_value = []
        for cell in row[0:3]:
            row_value.append(cell.value)
        L3.append(row_value)
    print(L3)

    #DATA INSERT
    bt = "O"
    btp = "  "
    if int(L1[0]) == 2 : bt, btp = btp, bt

    pddata = {'x1' : bt, 'x2' : btp, 'const' : L1[6],
              'a1' : L1[1],
              'a2' : L1[2],
              'a3' : L1[3],
              'a4' : L1[4],
              'a5' : L2[0],
              'a6' : L2[1],
              'b1' : L2[2].date(),
              'b2' : L2[3].date(),
              'd1' : L2[4].year,
              'd2' : L2[4].month,
              'd3' : L2[4].day
              }
    ppdata = L3
    print(L2[3].date())
    return pddata, ppdata
load_data("ddd")

#####

tmp = [['테스트', 1234567, 765544321], ['테스트1', 21234567, 1765544321]]

idx = 1
pg = 1
fdata = []

f = []
d = []#

for i in tmp : #ppdata
    if idx > 4 :
        pg = pg + 1
        idx = 1

    #f.append("C"+str(idx)+"1")
    #d.append("CONST") #pddata'const'

    fdata.append(("C"+str(idx)+"1", pg, "CONST"))#pddata'const'

    for j in range(2,5) :
        #f.append("C" + str(idx) + str(j))
        #d.append(i[j-2])
        fdata.append(("C" + str(idx) + str(j), pg, i[j-2]))

    idx = idx + 1