import os
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import tkinter.messagebox
from tkinter.filedialog import askopenfilename
## input typhoon name and make file
root = Tk()
root.title("Giai Ma Typh")
mainframe = ttk.Frame(root, padding = "60 40 60 40")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

ty_name = StringVar()
obs_year = StringVar()
obs_month = StringVar()

ty_name_entry = ttk.Entry(mainframe, width=10, textvariable= ty_name)
ty_name_entry.grid(column=1, row=1, sticky= W)
ttk.Label(mainframe, text="Typhoon Name").grid(column=0, row=1, sticky=W)


obs_year_entry = ttk.Entry(mainframe, width=10, textvariable= obs_year)
obs_year_entry.grid(column=1, row=2, sticky= W)
ttk.Label(mainframe, text="Year").grid(column=0, row=2, sticky=W)

obs_month_entry = ttk.Entry(mainframe, width=10, textvariable= obs_month)
obs_month_entry.grid(column=1, row=3, sticky= W)
ttk.Label(mainframe, text="Month").grid(column=0, row=3, sticky=W)

def get_input():
   global input_file
   input_file = askopenfilename()
   return input_file

ttk.Button(mainframe, text="Input", command = get_input).\
                      grid(column=3, row=1, sticky=W)
def pop_up():
   name = ty_name_entry.get()
   year = obs_year_entry.get()
   month = obs_month_entry.get()
   messagebox.showinfo("input complete\n", "typhoon name: "+name+"\n"+\
                       "Year: "+year+"\n"+ "Month: "+month+"\n"+\
                       "File_name: "+input_file)
   return
def print_file():
   name = ty_name_entry.get()
   year = obs_year_entry.get()
   month = obs_month_entry.get()
   print (name,year,month)
   print (input_file)
   return
ttk.Button(mainframe, text="show", command = pop_up).\
                      grid(column=3, row=2, sticky=W)
#ty_name = input("Enter typhoon name: ")
#cwd = os.getcwd()
#output_file = os.path.join(cwd,str(ty_name) + ".csv")
root.mainloop()
