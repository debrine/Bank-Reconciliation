from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from csv import reader
from openpyxl import load_workbook


#global variables
bank_file_path = None
sage_file_path = None


def compareLists(bank_statement, sage_statement):
    pass

def reconcile():
    if (bank_file_path == None or sage_file_path == None):
        pass
    else:   
        CSV = processCSV()
        excel = processExcel()

        for entry in CSV:
            print(entry["cheque_num"])

#found at https://stackoverflow.com/questions/8270092/remove-all-whitespace-in-a-string
def removeExtraSpaces(string):
    return(" ".join(string.split()))

def processCSV():
    formattedCSV = []
    with open(bank_file_path, newline='') as csvfile:
        csv_reader = reader(csvfile)
        for row in csv_reader:
            entry = {}
            entry["date"] = row[1]
            entry["comment"] = removeExtraSpaces(row[2])
            entry["source_num"] = row[3]
            entry["credit"] = row[4]
            entry["debit"] = row[5]

            formattedCSV.append(entry)
    return formattedCSV
    
    
def processExcel():
    workbook = load_workbook(filename=sage_file_path, read_only=True)
    sheet = workbook['Sheet1']
    rows = list(sheet.rows)
    rows = rows[5:-2]
    formattedExcel = []
    for row in rows:
        data = []
        for cell in row:
            data.append(cell.value)
        entry = {}

        entry["date"] = data[2]
        entry["comment"] = data[3]
        entry["source_num"] = data[4]
        entry["debit"] = data[6]
        entry["credit"] = data[7]

        formattedExcel.append(entry)
    return formattedExcel

        

def select_bank_file():
    global bank_file_path
    bank_file_path = filedialog.askopenfilename(initialdir = "/",title = "Select Bank File",filetypes = ( ("csv files","*.csv"), ))
    bank_file_name.set(bank_file_path.split('/')[-1])

def select_sage_file():
    global sage_file_path
    sage_file_path = filedialog.askopenfilename(initialdir = "/",title = "Select Sage File",filetypes = ( ("xlsx files","*.xlsx"), ))
    sage_file_name.set(sage_file_path.split('/')[-1])
    



root = Tk()
root.title("Bank Reconciliation")

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

bank_file_name = StringVar()
bank_file_name.set('None')

sage_file_name = StringVar()
sage_file_name.set('None')

ttk.Label(mainframe, text='Select Bank File:').grid(column=0, row=0, sticky=W, padx=(50, 15), pady=5)
ttk.Button(mainframe, text='Choose .csv File', command=select_bank_file).grid(column=1, row=0, padx=(15,50), pady=5)
ttk.Label(mainframe, text='Selected File:').grid(column=0, row=1, padx=(50, 15), pady=5, sticky=W)
ttk.Label(mainframe, textvariable=bank_file_name).grid(column=1, row=1, padx = 15, pady=5, sticky=W)

ttk.Label(mainframe, text='Select Sage File:').grid(column=2, row=0, sticky=W, padx=(50, 15), pady=5)
ttk.Button(mainframe, text='Choose .xlsx File', command=select_sage_file).grid(column=3, row=0, padx=(15, 50), pady=5)
ttk.Label(mainframe, text='Selected File:').grid(column=2, row=1, padx=(50,15), pady=5, sticky=W)
ttk.Label(mainframe, textvariable=sage_file_name).grid(column=3, row=1, padx=(15,50), pady=5, sticky=W)

ttk.Button(mainframe, text='Reconcile', command=reconcile).grid(column=4, row=0, rowspan=2, padx=(50,15), pady=15)


root.mainloop()