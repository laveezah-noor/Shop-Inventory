#imports
from tkinter import *
import os
import json

class Register:
    def __init__(self, parent):
        temp_name = StringVar()
        temp_email = StringVar()
        # temp_gender = StringVar()
        temp_password = StringVar()

        # Register Screen
        register_screen = Toplevel(parent)
        register_screen.title('Register')

        #Labels
        Label(register_screen, text="Please enter your details below to register", font=('Calibri',12)).grid(row=0,sticky=N,pady=10)
        Label(register_screen, text="Name", font=('Calibri',12)).grid(row=1,sticky=W)
        Label(register_screen, text="Email", font=('Calibri',12)).grid(row=2,sticky=W)
        # Label(register_screen, text="Gender", font=('Calibri',12)).grid(row=3,sticky=W)
        Label(register_screen, text="Password", font=('Calibri',12)).grid(row=4,sticky=W)
        notif = Label(register_screen, font=('Calibri',12))
        notif.grid(row=6,sticky=N,pady=10)

        #Entries
        Entry(register_screen,textvariable=temp_name).grid(row=1,column=0)
        Entry(register_screen,textvariable=temp_email).grid(row=2,column=0)
        # Entry(register_screen,textvariable=temp_gender).grid(row=3,column=0)
        Entry(register_screen,textvariable=temp_password,show="*").grid(row=4,column=0)

        def finish_reg():
            name = temp_name.get()
            email = temp_email.get()
            # gender = temp_gender.get()
            password = temp_password.get()
            path = "Users.json"
            all_accounts_file = open(path, 'r')
            all_accounts = (json.load(all_accounts_file)["Users"])
            print(all_accounts)

            if name == "" or email == "" or password == "":
                notif.config(fg="red", text="All fields required * ", font=('Ariel', 8))
            else:
                notif.config(text="", font=('Ariel', 8))

            i = 0
            while len(all_accounts) > i:
                for user in all_accounts:
                    print('User', user["name"])
                    if name == user["name"]:
                        notif.config(fg="red", font=('Calibri', 8), text="Account already exists")
                        all_accounts_file.close()
                        break
                    else:
                        i += 1
            else:
                user = {"name": name, "email": email, "password": password}
                all_accounts.append(user)
                all_accounts = json.dumps(all_accounts)
                print(all_accounts)
                user_file = open("Users.json", "w")
                user_file.write('{"Users":'+str(all_accounts)+'}')
                user_file.close()
                notif.config(fg='green', text="Account has been created")
                register_screen.destroy()

        #Buttons
        Button(register_screen, text="Register", command=finish_reg, font=('Calibri',12)).grid(row=5,sticky=N,pady=10)
