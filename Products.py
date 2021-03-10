from tkinter import *
from tkinter import ttk
from openpyxl import *
from tkinter import messagebox


class Products(Frame):
    def __init__(self, *args,**kwargs):
        Frame.__init__(self,*args,**kwargs)

        self.main = None
        self.file = ''
        workbook = load_workbook('data.xlsx')
        sheet = workbook.active
        i = 0
        while i < (sheet.max_row):
            print(i, sheet.max_row)
            if (sheet.cell(row=i+1,column=5).value) == 1:
                self.file = sheet.cell(row=i+1,column=4).value
                self.main = load_workbook(filename=self.file)
                self.main.active = self.main['Products']
                print(self.main.active)
                break
            i += 1
        else:
            messagebox.showerror('Error','No User')


        self.label = Label(self, text="Products",
                        font=("times new roman", 20, "bold"))
        self.label.grid(row=0, column=1)

        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.productprice = StringVar()
        self.productname = StringVar()
        self.productqty = StringVar()


        Manage_Frame = Frame(self, bd=4, relief=RIDGE, bg="crimson")
        Manage_Frame.place(x=20, y=100, width=450, height=600)

        m_title = Label(Manage_Frame, text="Manage Products", bg="crimson", fg="white",
                        font=("times new roman", 20, "bold"))
        m_title.grid(row=0, columnspan=2, pady=20)

        lbl_productname= Label(Manage_Frame, text="Product Name", bg="crimson", fg="white", font=("times new roman", 18, "bold"))
        lbl_productname.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        txt_productname= Entry(Manage_Frame, textvariable=self.productname, font=("times new roman", 18, "bold"), bd=5,
                         relief=GROOVE)
        txt_productname.grid(row=1, column=1, padx=10, pady=10, sticky="w")

        lbl_productprice= Label(Manage_Frame, text="Product Price", bg="crimson", fg="white", font=("times new roman", 18, "bold"))
        lbl_productprice.grid(row=2, column=0, padx=10, pady=10, sticky="w")
        txt_productprice= Entry(Manage_Frame, textvariable=self.productprice, font=("times new roman", 18, "bold"), bd=5,
                         relief=GROOVE)
        txt_productprice.grid(row=2, column=1, padx=10, pady=10, sticky="w")

        lbl_productqty= Label(Manage_Frame, text="Quantity", bg="crimson", fg="white", font=("times new roman", 18, "bold"))
        lbl_productqty.grid(row=3, column=0, padx=10, pady=10, sticky="w")
        txt_productqty= Entry(Manage_Frame, textvariable=self.productqty, font=("times new roman", 18, "bold"), bd=5,
                         relief=GROOVE)
        txt_productqty.grid(row=3, column=1, padx=10, pady=10, sticky="w")

        lbl_add_button = Button(Manage_Frame, text="Add", bg="crimson", fg="white",
                            font=("times new roman", 18, "bold"), command=self.add_data)
        lbl_add_button.grid(row=8, column=0, padx=10, pady=10, sticky="w")

        lbl_clear_button = Button(Manage_Frame, text="Clear", bg="crimson", fg="white",
                            font=("times new roman", 18, "bold"), command=self.clear)
        lbl_clear_button.grid(row=8, column=1, padx=10, pady=10, sticky="w")

        Detail_Frame = Frame(self, bd=4, relief=RIDGE, bg="crimson")
        Detail_Frame.place(x=500, y=100, width=750, height=580)

        btn_show = Button(Detail_Frame, text="Show Data", font=("times new roman", 13), bd=5,
                            relief=GROOVE, command=self.show_data, bg="crimson")
        btn_show.grid(row=0, column=2, padx=5, pady=5, sticky="w")

        # Table Frame
        Table_Frame = Frame(Detail_Frame, bd=4, relief=RIDGE, bg="crimson")
        Table_Frame.place(x=20, y=60)

        scroll_x = Scrollbar(Table_Frame, orient=HORIZONTAL)
        scroll_y = Scrollbar(Table_Frame, orient=VERTICAL)

        self.Product_table = ttk.Treeview(Table_Frame,
           columns=("productname", "productprice", "productqty"),
           xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set)
        scroll_x.pack(side=BOTTOM, fill=X)
        scroll_y.pack(side=RIGHT, fill=Y)
        scroll_x.config(command=self.Product_table.xview)
        scroll_y.config(command=self.Product_table.yview)
        # self.Product_table.heading("productid", text="Product#")
        self.Product_table.heading("productname", text="Name")
        self.Product_table.heading("productprice", text="Price")
        self.Product_table.heading("productqty", text="Quantity")
        self.Product_table['show'] = 'headings'  # removing extra index col at begining

        # setting up widths of cols
        # self.Product_table.column("productid", width=100)
        self.Product_table.column("productname", width=100)
        self.Product_table.column("productprice", width=100)
        self.Product_table.column("productqty", width=100)
        self.Product_table.pack(fill=BOTH, expand=1)  # fill both is used to fill cols around the frame
        # self.Product_table.bind("<ButtonRelease-1>", self.get_cursor)  # this is an event to select row

    def clear(self):
        self.productqty.set("")
        self.productname.set("")
        self.productprice.set("")

    def show_data(self):
        sheet = self.main.active
        row_max = sheet.max_row
        for i in self.Product_table.get_children():
            self.Product_table.delete(i)
        for i in range(row_max):
            # no = sheet.cell(column=1, row=i+1).value
            name = sheet.cell(column=1, row=i+1).value
            price = sheet.cell(column=2, row=i+1).value
            qty = sheet.cell(column=3, row=i+1).value
            print(name, price, qty)
            self.Product_table.insert('', 'end', text="", values=(name, price, qty))
        messagebox.showinfo("successfull", "Record has been updated.")

    def add_data(self):
        sheet = self.main.active
        print(sheet)
        name = self.productname.get()
        price = self.productprice.get()
        qty = self.productqty.get()
        print(name, price, qty)
        data = [name, price, qty]
        sheet.append(data)
        workbook.save(filename=self.file)
