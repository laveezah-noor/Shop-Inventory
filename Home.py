from tkinter import *
from tkinter import ttk


class Home(Frame):

    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        # self.title('Shop Inventory')
        # master.geometry('400x400')

        label = Label(self, text="Start Page")
        label.pack(padx=10, pady=10)
        page_one = Button(self, text="Page One", command=lambda:controller.show_frame(PageOne))
        page_one.pack()
        page_two = Button(self, text="Page Two", command=lambda:controller.show_frame(PageTwo))
        page_two.pack()
        # mainframe = ttk.Frame(master, padding="3 3 12 12")
        # mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
        # master.columnconfigure(0, weight=1)
        # master.rowconfigure(0, weight=1)
        #
        # ttk.Button(mainframe, text="Get Started", command=None).grid(column=3, row=3, sticky=(N, W, E, S))
        # ttk.Label(mainframe, text="Shop Inventory").grid(column=3, row=1, sticky=W)
        # #
        # for child in mainframe.winfo_children():
        #     child.grid_configure(padx=15, pady=15)

        # feet_entry.focus()

        # master.bind("<Return>", self.calculate)


# root = Tk()
# Home(root)
# root.mainloop()
