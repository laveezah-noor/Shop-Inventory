from openpyxl import load_workbook
# from openpyxl import Workbook
# wb = Workbook()
workbook = load_workbook(filename="data.xlsx")
# workbook = Workbook()

sheet = workbook.active
sheet1 = workbook['Users']
# sheet["A1"] = "hello"
# sheet["B1"] = "world!"
# sheet.title =
# products_sheet = workbook["Products"]

# workbook.save(filename="data.xlsx")
# print(workbook.sheetnames)
# print(workbook.active)

print(workbook.active.title)
