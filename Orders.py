from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from openpyxl import *
import datetime


class Orders(Frame):
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
                self.main.active = self.main['Orders']
                print(self.main.active)
                break

            i += 1
        else:
            messagebox.showerror('Error','No User')

        self.amount = StringVar()
        self.product = StringVar()
        self.customer = StringVar()
        self.qty = StringVar()
       
        Manage_Frame = Frame(self, bd=4, relief=RIDGE, bg="crimson")
        Manage_Frame.place(x=20, y=100, width=450, height=600)

        m_title = Label(Manage_Frame, text="Manage Orders", bg="crimson", fg="white",
                        font=("times new roman", 20, "bold"))
        m_title.grid(row=0, columnspan=2, pady=20)

        lbl_amount = Label(Manage_Frame, text="Amount", bg="crimson", fg="white", font=("times new roman", 18, "bold"))
        lbl_amount.grid(row=1, column=0, padx=20, pady=10, sticky="w")
        txt_amount = Entry(Manage_Frame, textvariable=self.amount, font=("times new roman", 18, "bold"), bd=5,
                         relief=GROOVE)
        txt_amount.grid(row=1, column=1, padx=20, pady=10, sticky="w")

        lbl_product = Label(Manage_Frame, text="Product", bg="crimson", fg="white", font=("times new roman", 18, "bold"))
        lbl_product.grid(row=2, column=0, padx=20, pady=10, sticky="w")
        combo_product = ttk.Combobox(Manage_Frame, textvariable=self.product, font=("times new roman", 13, "bold"),
                                    state="readonly")
        combo_product['values'] = ("Male", "Female", "Other")
        combo_product.grid(row=2, column=1, padx=20, pady=10, sticky="w")

        lbl_customer = Label(Manage_Frame, text="Customer", bg="crimson", fg="white", font=("times new roman", 18, "bold"))
        lbl_customer.grid(row=3, column=0, padx=20, pady=10, sticky="w")
        combo_customer = ttk.Combobox(Manage_Frame, textvariable=self.customer, font=("times new roman", 13, "bold"),
                                    state="readonly")
        combo_customer['values'] = ("Male", "Female", "Other")
        combo_customer.grid(row=3, column=1, padx=20, pady=10, sticky="w")

        lbl_qty = Label(Manage_Frame, text="Quantity", bg="crimson", fg="white",
                            font=("times new roman", 18, "bold"))
        lbl_qty.grid(row=5, column=0, padx=20, pady=10, sticky="w")
        txt_qty = Entry(Manage_Frame, textvariable=self.qty, font=("times new roman", 18, "bold"), bd=5,
                            relief=GROOVE)
        txt_qty.grid(row=5, column=1, padx=20, pady=10, sticky="w")

        lbl_add_button = Button(Manage_Frame, text="Add", bg="crimson", fg="white",
                            font=("times new roman", 18, "bold"), command=self.add_data)
        lbl_add_button.grid(row=8, column=0, padx=20, pady=10, sticky="w")
        
        lbl_clear_button = Button(Manage_Frame, text="Clear", bg="crimson", fg="white",
                            font=("times new roman", 18, "bold"), command=self.clear)
        lbl_clear_button.grid(row=8, column=1, padx=20, pady=10, sticky="w")
         
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
 
        self.Order_table = ttk.Treeview(Table_Frame,
           columns=("date", "customer", "product", "amount", "qty"),
           xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set)
        scroll_x.pack(side=BOTTOM, fill=X)
        scroll_y.pack(side=RIGHT, fill=Y)
        scroll_x.config(command=self.Order_table.xview)
        scroll_y.config(command=self.Order_table.yview)
        # self.Order_table.heading("order", text="Order#")
        self.Order_table.heading("date", text="Date")
        self.Order_table.heading("customer", text="Customer Name")
        self.Order_table.heading("product", text="Product Name")
        self.Order_table.heading("amount", text="Amount")
        self.Order_table.heading("qty", text="Quantity")
        self.Order_table['show'] = 'headings'  # removing extra index col at begining
 
        # setting up widths of cols
        # self.Order_table.column("order", width=100)
        self.Order_table.column("date", width=100)
        self.Order_table.column("customer", width=100)
        self.Order_table.column("product", width=100)
        self.Order_table.column("amount", width=100)
        self.Order_table.column("qty", width=100)
        self.Order_table.pack(fill=BOTH, expand=1)  # fill both is used to fill cols around the frame
        # self.Order_table.bind("<ButtonRelease-1>", self.get_cursor)  # this is an event to select row
 
    def clear(self):
        self.amount.set("")
        self.customer.set("")
        self.product.set("")
        self.qty.set("")


    def show_data(self):
        sheet = self.main.active
        row_max = sheet.max_row
        col_max = sheet.max_column
        for i in self.Order_table.get_children():
            self.Order_table.delete(i)
        for i in range(1,row_max):
            # order = sheet.cell(column=1, row=i+1).value
            date = sheet.cell(column=1, row=i+1).value
            customer = sheet.cell(column=2, row=i+1).value
            product = sheet.cell(column=3, row=i+1).value
            amount = sheet.cell(column=4, row=i+1).value
            qty = sheet.cell(column=5, row=i+1).value
            print(date, customer, product, amount, qty)
            self.Order_table.insert('', 'end', text="", values=(date, customer, product, amount, qty))
        messagebox.showinfo("successfull", "Record has been updated.")

    def add_data(self):
        sheet = self.main.active
        print(sheet)
        date = datetime.datetime.now().strftime("%x")
        customer = self.customer.get()
        product = self.product.get()
        amount = self.amount.get()
        qty = self.qty.get()
        print(date, customer, product, amount, qty)
        data = [date, customer, product, amount, qty]
        sheet.append(data)
        self.main.save(filename=self.file)
