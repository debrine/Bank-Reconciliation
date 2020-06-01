from tkinter import *
from tkinter import ttk
from tkinter import filedialog


bank_file_name = None

sage_file_name = None


def compareLists(bank_statement, sage_statement):
    pass

def reconcile():
    pass

def select_bank_file():
    global bank_file_name
    bank_file_name = filedialog.askopenfilename(initialdir = "/",title = "Select Bank File",filetypes = ( ("csv files","*.csv"), ))
    

def select_sage_file():
    global sage_file_name
    sage_file_name = filedialog.askopenfilename(initialdir = "/",title = "Select Sage File",filetypes = ( ("xlsx files","*.xlsx"), ))
    
root = Tk()
root.title("Bank Reconciliation")

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

ttk.Label(mainframe, text='Select Bank File:').grid(column=0, row=0, sticky=W)
ttk.Button(mainframe, text='Choose .csv File', command=select_bank_file).grid(column=0, row=1)

ttk.Label(mainframe, text='Select Sage File:').grid(column=1, row=0, sticky=W)
ttk.Button(mainframe, text='Choose .xlsx File', command=select_sage_file).grid(column=1, row=1)

ttk.Button(mainframe, text='Reconcile', command=reconcile).grid(column=2, row=0, rowspan=2)

for child in mainframe.winfo_children(): child.grid_configure(padx=50, pady=15)

root.mainloop()