def selectItemtoEdit(event):
    Start_flag = True
    global row_to_change
    global col_to_change
    global value_to_change
    region = treeview.identify_region(event.x,event.y)
    if region not in ("tree","cell"):
        return

    column = treeview.identify_column(event.x)
    column_index=int(column[1:])-1
    iid=treeview.focus()
    iid_index = int(iid[1:])-1
    selected_value=treeview.item(iid)
    selected_text=selected_value.get("values")[column_index]

    #print(selected_text)

    column_box=treeview.bbox(iid,column)
    entry_edit= tk.Entry(result)

    entry_edit.editing_column_index=column_index
    entry_edit.editing_row_index=iid_index

    #Click other area to save changes
    treeview.bind("<Button-1>", on_focus_out)


