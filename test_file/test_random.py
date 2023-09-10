from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import tkinter.messagebox


top = Tk()
L1 = Label(top, text="User Name")
L1.pack( side = LEFT)
E1 = Entry(top, bd =5)
E1.pack(side = RIGHT)
E1.get()
def hello():
   messagebox.showinfo("Say Hello",E1)

B1 = ttk.Button(top, text = "Hello ", command = hello)
B1.pack()
top.mainloop()
