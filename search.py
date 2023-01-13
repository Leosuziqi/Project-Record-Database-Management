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
#Search Function
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
    display()

    other_fucntions()

    re_search=ttk.Button(result, text="Return to Search", command=back_to_search)
    re_search.place(relx=0.6,rely=0.855,width=100,height=40)
    exit_2 = ttk.Button(result, text="Close", command=search_close)
    exit_2.place(relx=0.8, rely=0.855, width=100, height=40)


    # Add Buttons to Edit and Delete the Treeview items
    treeview.bind('<Double-1>', selectItemtoEdit)
    edit_btn = ttk.Button(result, text="Edit", command=edit)
    edit_btn.place(relx=0.2,rely=0.855,width=100,height=40)
    del_btn = ttk.Button(result, text="Delete", command=delete)
    del_btn.place(relx=0.4,rely=0.855,width=100,height=40)

