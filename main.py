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
#cols = []

#frame1 = LabelFrame(main,text = 'Search information:')
#frame1.pack(expand = 'yes', fill = 'both')

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
def search():
    #search
    rows = []
    id_flag = FALSE
    Company_flag = FALSE
    Location_flag = FALSE
    if ProjectNum.get():
        search_id=ProjectNum.get()
        id_flag = TRUE

    if CompanyName.get():
        search_Company=CompanyName.get()
        Company_flag = TRUE

    if Location.get():
        search_Location=Location.get()
        Location_flag = TRUE

    ProjectNum.configure(state=tk.NORMAL)
    CompanyName.configure(state=tk.NORMAL)
    Location.configure(state=tk.NORMAL)
    #email.configure(state=tk.NORMAL)

    file = openpyxl.load_workbook(excel_path, data_only=True)
    sheet=file['Freezer Job Record']

    #Save every matched row into cols, then append all cols to rows[] as the searched result
    cols=[]
    itr=2   #start searching from row 2
    for cell in sheet.iter_rows(min_row=2,max_row=sheet.max_row,min_col=1,max_col=22,values_only=FALSE):    #Iteration in rows
        #values_only?
        if not (id_flag|Company_flag|Location_flag):
            continue
        #check if one or two or three inserted data matched
        if (id_flag and (str(search_id).lower() in str(sheet[itr][0].value).lower())) or id_flag == FALSE:
            #print(id_flag)
            #print(sheet[itr][0].value)
            if (Company_flag and (str(search_Company).lower() in str(sheet[itr][1].value).lower())) or Company_flag == FALSE:
                if (Location_flag and str(search_Location).lower() in str(sheet[itr][2].value).lower()) or Location_flag == FALSE:
                    for index in cell:
                        #print(index)
                        cols.append(index.value)
                    rows.append(cols)
                    cols=[]
        itr = itr + 1

    #print(len(rows))

    #display
    main.withdraw()
    headers = []
    for header in sheet[1]:
        if len(headers)<22:
            headers.append(header)

    result = tk.Tk()
    result.title('Search Result')
    result.geometry('1100x400')



Button(main,text="Search",command=search).place(x=230, y=400)
# Button for closing
Button(main, text="Exit", command=main.destroy).place(x=460, y=400)


main.mainloop()
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
