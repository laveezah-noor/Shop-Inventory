#imports
from tkinter import *
import os


class Register:
    def __init__(self, parent):
        temp_name = StringVar()
        temp_age = StringVar()
        temp_gender = StringVar()
        temp_password = StringVar()

        # Register Screen
        register_screen = Toplevel(parent)
        register_screen.title('Register')

        #Labels
        Label(register_screen, text="Please enter your details below to register", font=('Calibri',12)).grid(row=0,sticky=N,pady=10)
        Label(register_screen, text="Name", font=('Calibri',12)).grid(row=1,sticky=W)
        Label(register_screen, text="Age", font=('Calibri',12)).grid(row=2,sticky=W)
        Label(register_screen, text="Gender", font=('Calibri',12)).grid(row=3,sticky=W)
        Label(register_screen, text="Password", font=('Calibri',12)).grid(row=4,sticky=W)
        notif = Label(register_screen, font=('Calibri',12))
        notif.grid(row=6,sticky=N,pady=10)

        #Entries
        Entry(register_screen,textvariable=temp_name).grid(row=1,column=0)
        Entry(register_screen,textvariable=temp_age).grid(row=2,column=0)
        Entry(register_screen,textvariable=temp_gender).grid(row=3,column=0)
        Entry(register_screen,textvariable=temp_password,show="*").grid(row=4,column=0)

        def finish_reg():
            name = temp_name.get()
            age = temp_age.get()
            gender = temp_gender.get()
            password = temp_password.get()
            all_accounts = os.listdir('Users.txt')
            print(all_accounts)

            # if name == "" or age == "" or gender == "" or password == "":
            #     notif.config(fg="red",text="All fields requried * ")
            #     return
            #
            # for name_check in all_accounts:
            #     if name == name_check:
            #         notif.config(fg="red",text="Account already exists")
            #         return
            #     else:
            #         user_file = open("./Users.txt","w")
            #         user_file.write(name+'\n')
            #         user_file.write(password+'\n')
            #         user_file.write(age+'\n')
            #         user_file.write(gender+'\n')
            #         user_file.close()
            #         notif.config(fg="green", text="Account has been created")

        #Buttons
        Button(register_screen, text="Register", command=finish_reg, font=('Calibri',12)).grid(row=5,sticky=N,pady=10)
