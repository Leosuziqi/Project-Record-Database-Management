# Data fetching from Excel files
import pathlib
# Author: Leo Su
# Data 19-Dec-2022

from tkinter import *
import tkinter as tk
import numpy as np
import openpyxl,xlrd
from tkinter import Tk, mainloop,TOP
from openpyxl import Workbook
#import cv2
from tkinter import ttk
from tkinter.messagebox import showinfo
from search import search
from fetch_calendar import *


main=Tk()
main.title('Search_Page')
main.geometry('800x500')
main.config(highlightbackground="black", highlightthickness=2)

"""
file = pathlib.Path("test.xlsx")
if file.exists():
    pass
else:
    #Erro Message: File not Existed.
    error_page=Tk()
    error_page.title('error')
    error_page.geometry("200*100")
    error_label = tk.Label(error_page, text="File not Existed.")
    error_label.place(anchor="center")
"""
excel_path = r".\Freezer Job Record.xlsx"
#excel_path = r"I:\Support Group\Service\Freezer Job Record.xlsx"

#frame1 = LabelFrame(main,text = 'Search information:')
#frame1.pack(expand = 'yes', fill = 'both')

output_table = fetch_calendar()

item1_name=Label(main,text="Project #:")
item1_name.place(x=50,y=70)
item2_name=Label(main,text="Company Name:")
item2_name.place(x=50,y=100)
item3_name=Label(main,text="Location:")
item3_name.place(x=50,y=130)

ProjectNum=Entry(main)
ProjectNum.place(x=250,y=70)
CompanyName=Entry(main)
CompanyName.place(x=250,y=100)
Location=Entry(main)
Location.place(x=250,y=130)

Button(main,text="Search",command=search).place(x=230, y=400)
# Button for closing
Button(main, text="Exit", command=main.destroy).place(x=460, y=400)


main.mainloop()
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
