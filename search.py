# Author: Leo Su
# Data 19-Dec-2022

import numpy as np
from tkinter import Tk, mainloop,TOP, ttk
from openpyxl import Workbook
from tkinter.messagebox import showinfo
import pathlib

from TreeviewAction import *
from fetch_calendar import *

#Search if records match with contents from keyboard entry
def search():
    rows = []
    id_flag = FALSE
    Company_flag = FALSE
    Location_flag = FALSE
    
    #check which entry gets input
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

    #Save every matched row into cols, then append all cols to rows[] as the searched result
    cols=[] 
    for itr, cell in sheet.iter_rows(min_row=2,max_row=sheet.max_row,min_col=1,max_col=22,values_only=FALSE):   
        
        # no input
        if not (id_flag|Company_flag|Location_flag):
            continue
            
        #check if one or two or three inserted data matched
        if (id_flag and (str(search_id).lower() in str(sheet[itr][0].value).lower())) or id_flag == FALSE:
            #print(id_flag)
            #print(sheet[itr][0].value)
            
            #check matchs
            if (Company_flag and (str(search_Company).lower() in str(sheet[itr][1].value).lower())) or Company_flag == FALSE:
                if (Location_flag and str(search_Location).lower() in str(sheet[itr][2].value).lower()) or Location_flag == FALSE:
                    # save mathed result to ROWS
                    for index in cell:
                        cols.append(index.value)
                    rows.append(cols)
                    cols=[]
        itr = itr + 1

    #print(len(rows))
    
    # display result with Treeview API
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

