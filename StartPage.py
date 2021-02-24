from tkinter import *
from tkinter import ttk
from Login import Login


class StartPage(Frame):

    def __init__(self, parent, controller):
        Frame.__init__(self, parent)

        label1 = Label(self, text="Let's Get Started!")
        label2 = Label(self, text="With Your Own Inventory")
        label1.pack(padx=10, pady=10)
        label2.pack(padx=10, pady=10)
        login = Button(self, text="Get Started", command=lambda:controller.show_frame(Login))
        login.pack()
