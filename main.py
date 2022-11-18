import os

import xlrd

data = xlrd.open_workbook("name_sheet.xls")
table = data.sheet_by_index(0)

for name in os.listdir("source"):
    print(name)
    for i in range(1, 45):
        if table.cell_value(i, 2) == name:
            img_name = os.listdir("source/" + name)[0]
            type_name = img_name.split(".")[len(img_name.split(".")) - 1]
            print(img_name)
            os.rename("source/" + name + "/" + img_name, "target/" + name + str(int(table.cell_value(i, 1))) + "." + type_name)
            break
