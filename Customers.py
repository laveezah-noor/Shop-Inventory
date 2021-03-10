from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from openpyxl import *


class Customers(Frame):
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
                self.main.active = self.main['Customers']
                print(self.main.active)
                break
            i += 1
        else:
            messagebox.showerror('Error','No User')


        self.label = Label(self, text="Hi This is Tab2")
        self.label.grid(row=1,column=0,padx=10,pady=10)  
        
        self.contact_var = StringVar()
        self.name_var = StringVar()
        self.email_var = StringVar()
       
        Manage_Frame = Frame(self, bd=4, relief=RIDGE, bg="pink")
        Manage_Frame.place(x=10, y=50, width=450, height=600)

        m_title = Label(Manage_Frame, text="Manage Customers", bg="pink", fg="white",
                        font=("times new roman", 20, "bold"))
        m_title.grid(row=0, columnspan=2, pady=10)

        lbl_name = Label(Manage_Frame, text="Name", bg="crimson", fg="white", font=("times new roman", 18, "bold"))
        lbl_name.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        txt_name = Entry(Manage_Frame, textvariable=self.name_var, font=("times new roman", 18, "bold"), bd=5,
                         relief=GROOVE)
        txt_name.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        lbl_email = Label(Manage_Frame, text="Email", bg="crimson", fg="white", font=("times new roman", 18, "bold"))
        lbl_email.grid(row=2, column=0, padx=10, pady=5, sticky="w")
        txt_email = Entry(Manage_Frame, textvariable=self.email_var, font=("times new roman", 18, "bold"), bd=5,
                         relief=GROOVE)
        txt_email.grid(row=2, column=1, padx=10, pady=5, sticky="w")

        lbl_contact = Label(Manage_Frame, text="Contact", bg="crimson", fg="white", font=("times new roman", 18, "bold"))
        lbl_contact.grid(row=3, column=0, padx=10, pady=5, sticky="w")
        txt_contact = Entry(Manage_Frame, textvariable=self.contact_var, font=("times new roman", 18, "bold"), bd=5,
                          relief=GROOVE)
        txt_contact.grid(row=3, column=1, padx=10, pady=5, sticky="w")

        lbl_add_button = Button(Manage_Frame, text="Add", bg="pink", fg="white",
                            font=("times new roman", 18, "bold"), command=self.add_data)
        lbl_add_button.grid(row=8, column=0, padx=10, pady=5, sticky="w")
        
        lbl_clear_button = Button(Manage_Frame, text="Clear", bg="pink", fg="white",
                            font=("times new roman", 18, "bold"), command=self.clear)
        lbl_clear_button.grid(row=8, column=1, padx=10, pady=5, sticky="w")
         
        Detail_Frame = Frame(self, bd=4, relief=RIDGE, bg="pink")
        Detail_Frame.place(x=250, y=50, width=750, height=580)
 
        btn_show = Button(Detail_Frame, text="Show Data", font=("times new roman", 13), bd=5,
                            relief=GROOVE, command=self.show_data, bg="pink")
        btn_show.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        
        # Table Frame
        Table_Frame = Frame(Detail_Frame, bd=4, relief=RIDGE, bg="pink")
        Table_Frame.place(x=20, y=60)
 
        scroll_x = Scrollbar(Table_Frame, orient=HORIZONTAL)
        scroll_y = Scrollbar(Table_Frame, orient=VERTICAL)
 
        self.Customer_table = ttk.Treeview(Table_Frame,
           columns=("name", "email", "contact", "orders"),
           xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set)
        scroll_x.pack(side=BOTTOM, fill=X)
        scroll_y.pack(side=RIGHT, fill=Y)
        scroll_x.config(command=self.Customer_table.xview)
        scroll_y.config(command=self.Customer_table.yview)
        # self.Customer_table.heading("no", text="ID#")
        self.Customer_table.heading("name", text="Name")
        self.Customer_table.heading("email", text="Email Address")
        self.Customer_table.heading("contact", text="Contact")
        self.Customer_table.heading("orders", text="Orders")
        self.Customer_table['show'] = 'headings'  # removing extra index col at begining
 
        # setting up widths of cols
        # self.Customer_table.column("no", width=100)
        self.Customer_table.column("name", width=100)
        self.Customer_table.column("email", width=100)
        self.Customer_table.column("contact", width=100)
        self.Customer_table.column("orders", width=100)
        self.Customer_table.pack(fill=BOTH, expand=1)  # fill both is used to fill cols around the frame
        # self.Customer_table.bind("<ButtonRelease-1>", self.get_cursor)  # this is an event to select row
 
    def clear(self):
        self.name_var.set("")
        self.email_var.set("")
        self.contact_var.set("")

    def show_data(self):
        sheet = self.main.active
        row_max = sheet.max_row
        for i in self.Customer_table.get_children():
            self.Customer_table.delete(i)
        for i in range(row_max):
            # no = sheet.cell(column=1, row=i+1).value
            name = sheet.cell(column=1, row=i+1).value
            email = sheet.cell(column=2, row=i+1).value
            contact = sheet.cell(column=3, row=i+1).value
            orders = sheet.cell(column=4, row=i+1).value
            print(name, email, contact, orders)
            self.Customer_table.insert('', 'end', text="", values=( name, email, contact, orders))
        messagebox.showinfo("successfull", "Record has been updated.")

    def add_data(self):
        sheet = self.main.active
        print(sheet)
        name = self.name_var.get()
        email = self.email_var.get()
        contact = self.contact_var.get()
        order = 0
        print(name, email, contact, order)
        data = [name, email, contact, order]
        sheet.append(data)
        workbook.save(filename='data.xlsx')
