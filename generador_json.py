from openpyxl import Workbook
from openpyxl import load_workbook
import json

wb = load_workbook("BD_EXCEL.xlsx")
data = []
for sheet_title in wb.get_sheet_names():
    sheet = wb.get_sheet_by_name(sheet_title)
    for row in sheet.rows:
        fields = {}
        pk = 0
        if row[0].value is None:
            break
        for cell in row:
            if cell.column == 1:
                pk = cell.value
            if cell.value is None:
                fields[sheet.cell(row=1, column=cell.column).value] = ""
            else:
                fields[sheet.cell(row=1, column=cell.column).value] = cell.value
        data.append({
            "model" : sheet_title,
            "pk" : pk,
            "fields" : fields
        })

with open('data.json', 'w') as file:
    json.dump(data, file, indent=4)






