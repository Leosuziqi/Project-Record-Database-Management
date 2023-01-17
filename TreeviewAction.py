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


def edit():
    # Get selected item to Edit
    prject_id_to_change=str(rows[row_to_change][0])
    print(prject_id_to_change)
    itr_row=2
    for cell_2 in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=22, values_only=FALSE):
        if prject_id_to_change in str(sheet[itr_row][0].value):
            sheet.cell(row=itr_row,column=col_to_change+1).value = value_to_change
            file.save(excel_path)
            break

        itr_row = itr_row + 1

    #update_file()
    result.destroy()
    search()

def delete():
    # Get selected item to Delete
    selected_item = treeview.selection()[0]
    treeview.delete(selected_item)


def update_file():
    # Get selected item to Delete
    selected_item = treeview.selection()[0]
    treeview.delete(selected_item)
