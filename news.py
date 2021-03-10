from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from openpyxl import *
from Orders import Orders
from Products import Products
from Customers import Customers


class App(Tk):
    def __init__(self,*args,**kwargs):
       Tk.__init__(self,*args,**kwargs)
       self.notebook = ttk.Notebook()
       self.add_tab()
       self.notebook.place(width=1300, height=700)
  
    def add_tab(self):
        tab1 = Products(self.notebook)
        tab2 = Customers(self.notebook) 
        tab3 = Orders(self.notebook) 
        self.notebook.add(tab1,text="Products")
        self.notebook.add(tab2,text="Customers")
        self.notebook.add(tab3,text="Orders")
  
  
my_app = App()
my_app.geometry("1350x700+0+0")
my_app.mainloop()