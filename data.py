from openpyxl import load_workbook, Workbook

# workbook = load_workbook(filename="data.xlsx")
workbook = Workbook()

sheet = workbook.active

# sheet["A1"] = "hello"
# sheet["B1"] = "world!"
sheet.title = 'Users'
# products_sheet = workbook["Products"]

workbook.save(filename="data.xlsx")
# print(workbook.sheetnames)
# print(workbook.active)

print(sheet.title)
