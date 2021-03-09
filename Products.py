from tkinter import *
from tkinter import ttk
from openpyxl import *


class Product(Frame):
    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        workbook = load_workbook('data.xlsx')
        print("Product==> ",workbook.active)
        sheet = workbook.active
        userWorkbook = load_workbook()
        print("Product==> ",userWorkbook.active)
        userSheet = userWorkbook.active

        


        self.productId = productId
        self.productPrice = productPrice
        self.productName = productName

    def addProduct(self, productId, productPrice, productName):
        self.userSheet.append([productId, productPrice, productName])

    def showProducts(self):
        row_max = self.userSheet.max_row
        col_max = self.userSheet.max_column
