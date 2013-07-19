## Program will create UPDATE, DELETE, or INSERT statements from one of:
##  - CSV File
##  - XLS File
##  - XLSX File

#file = open('../make-a-statement/

# Program failed to parse file that I edited in LibreOffice.

from Tkinter import *
from tkFileDialog import askopenfilename
from openpyxl.reader.excel import load_workbook
from openpyxl.cell import Cell
import sys

update = 'UPDATE '
uSet = ' SET '
insert = 'INSERT INTO '
delete = 'DELETE '
where = ' WHERE '
equals = ' = '
tick = '\''

def execute():
    sys.exit()

def execute_xlsx(xlsxFile):
    wb = load_workbook(xlsxFile)
    sheetNames = wb.get_sheet_names() # sheetNames is Array of sheet names
    sheets = []
    
    for i in range(0,len(sheetNames)):
        print sheetNames[i]
        print i
        sheets.append(wb.get_sheet_by_name(sheetNames[i]))

    # at this point we have all worksheets stored in the sheets Array

    for i in range(0,len(sheets)):
        for rIndex in range(0,len(sheets[i].rows)):
            for cIndex in range(0,len(sheets[i].columns)):
                # print update + sheetName[i] + uSet +
                print sheets[i].cell(row=rIndex,column=cIndex).value

    print sheets[0].cell('A1').value
    #print sheets[0].cell(row=1,column=1).value
        
    print sheets[0]
    print filename + ' is an xlsx'
    

def getFile():
    filename = askopenfilename()

    if filename[-5:] == '.xlsx':
        execute_xlsx(filename)
    
    fileLabel.grid(row=0,columnspan=4)
    go.grid(row=2,column=3)
    Radiobutton(root, text="UPDATE", variable=operation, value=1).grid(row=1,column=1)
    Radiobutton(root, text="INSERT", variable=operation, value=2).grid(row=2,column=1)
    Radiobutton(root, text="DELETE", variable=operation, value=3).grid(row=3,column=1)

filename = ''


root = Tk() # root widget, only 1 per program!
root.wm_title("Make a Statement")
f = Frame(root, height=100, width=300)
f.grid_propagate(0)
f.grid()

fileLabel = Label(root, text=filename)
go = Button(root, text="OK", command=execute)
loadFile = Button(root, text="Load File", command=getFile).grid(row=1,columnspan=1)
# WHEN WE WANT TO BULD EXIT:
# def callback():
#     sys.exit()

operation = 0 # Default 0 option is for SELECT


root.mainloop()


## main program

