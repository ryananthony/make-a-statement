## Program will create UPDATE, DELETE, or INSERT statements from one of:
##  - CSV File
##  - XLS File
##  - XLSX File

#file = open('../make-a-statement/

from Tkinter import *
from tkFileDialog import askopenfilename
from openpyxl.reader.excel import load_workbook
import sys


def execute():
    sys.exit()
    

def getFile():
    filename = askopenfilename()

    if filename[-5:] == '.xlsx':
        print filename + ' is an xlsx'
    
    fileLabel.grid(row=0,columnspan=4)
    go.grid(row=2,column=3)
    Radiobutton(root, text="UPDATE", variable=operation, value=1).grid(row=1,column=1)
    Radiobutton(root, text="INSERT", variable=operation, value=2).grid(row=2,column=1)
    Radiobutton(root, text="DELETE", variable=operation, value=3).grid(row=3,column=1)

filename = ''


root = Tk() # root widget, only 1 per program!
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

