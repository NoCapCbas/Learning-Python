from email.mime import base
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import pyodbc
import win32com.client as win32
from pywintypes import com_error
from pathlib import Path
import sys
import pandas as pd  # only used for synthetic data
import numpy as np  # only used for synthetic data
import random  # only used for synthetic data
from datetime import datetime  # only used for synthetic dataimport PIL
from PIL import ImageTk,Image
import os
win32c = win32.constants

fields = ('Excel File Name', 'Excel Sheet Name', 'Save Path', 'Server', 'SQL Code')
folder_path = ''
homePath = os.path.expanduser('~')
basePath = homePath + '\\Documents'

def create_excel_file(f_path: Path, f_name: str, sheet_name: str, server: str, sqlQuery: str):
    
    filename = f_path / f_name
    
    # connects to DB3 grabbing TDM data availability
    conn = pyodbc.connect(
                            'Driver={SQL Server};'
                            f'Server={server};'
                            'UID=sa;'
                            'PWD=Harpua88;'
                            'Trusted_Connection=No;'
    )
    
    df = pd.read_sql_query(sqlQuery, conn)

    # create the dataframe and save it to Excel
    df.to_excel(filename, index=False, sheet_name=sheet_name, float_format='%.2f')

def pivot_table(wb: object, ws1: object, pt_ws: object, ws_name: str, pt_name: str, pt_rows: list, pt_cols: list, pt_filters: list, pt_fields: list):
    """
    wb = workbook1 reference
    ws1 = worksheet1
    pt_ws = pivot table worksheet number
    ws_name = pivot table worksheet name
    pt_name = name given to pivot table
    pt_rows, pt_cols, pt_filters, pt_fields: values selected for filling the pivot tables
    """

    # pivot table location
    pt_loc = len(pt_filters) + 2
    
    # grab the pivot table source data
    pc = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=ws1.UsedRange)
    
    # create the pivot table object
    pc.CreatePivotTable(TableDestination=f'{ws_name}!R{pt_loc}C1', TableName=pt_name)

    # selecte the pivot table work sheet and location to create the pivot table
    pt_ws.Select()
    pt_ws.Cells(pt_loc, 1).Select()

    # Sets the rows, columns and filters of the pivot table
    for field_list, field_r in ((pt_filters, win32c.xlPageField), (pt_rows, win32c.xlRowField), (pt_cols, win32c.xlColumnField)):
        for i, value in enumerate(field_list):
            pt_ws.PivotTables(pt_name).PivotFields(value).Orientation = field_r
            pt_ws.PivotTables(pt_name).PivotFields(value).Position = i + 1

    # Sets the Values of the pivot table
    for field in pt_fields:
        pt_ws.PivotTables(pt_name).AddDataField(pt_ws.PivotTables(pt_name).PivotFields(field[0]), field[1], field[2]).NumberFormat = field[3]

    # Visiblity True or Valse
    pt_ws.PivotTables(pt_name).ShowValuesRow = True
    pt_ws.PivotTables(pt_name).ColumnGrand = True
def run_excel(f_path: Path, f_name: str, sheet_name: str):

    filename = f_path / f_name

    # create excel object
    excel = win32.gencache.EnsureDispatch('Excel.Application')

    # excel can be visible or not
    excel.Visible = True  # False
    
    # try except for file / path
    try:
        wb = excel.Workbooks.Open(filename)
    except com_error as e:
        if e.excepinfo[5] == -2146827284:
            print(f'Failed to open spreadsheet.  Invalid filename or location: {filename}')
        else:
            raise e
        sys.exit(1)

    # set worksheet
    ws1 = wb.Sheets(sheet_name)

    # Setup and call pivot_table
    ws2_name = 'CH'
    wb.Sheets.Add().Name = ws2_name
    ws2 = wb.Sheets(ws2_name)
    
    pt_name = 'CH'  # must be a string
    pt_rows = ['PERIOD']  # must be a list
    pt_cols = ['CH']  # must be a list
    pt_filters = ['CTY_RPT']  # must be a list
    # [0]: field name [1]: pivot table column name [3]: calulation method [4]: number format
    pt_fields = [  # must be a list of lists
                 ['NbrCommods', 'Sum of NbrCommods', win32c.xlSum, '#,###'],
                ]
    pivot_table(wb, ws1, ws2, ws2_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields)

    # Setup and call pivot_table
    ws2_name = 'VALUE'
    wb.Sheets.Add().Name = ws2_name
    ws2 = wb.Sheets(ws2_name)
    
    pt_name = 'VALUE'  # must be a string
    pt_rows = ['PERIOD']  # must be a list
    pt_cols = ['UNIT1']  # must be a list
    pt_filters = ['CTY_RPT', 'CH']  # must be a list
    # [0]: field name [1]: pivot table column name [3]: calulation method [4]: number format
    pt_fields = [  # must be a list of lists
                 ['USD', 'Sum of USD', win32c.xlSum, '$#,##0.00'],
                ]
    pivot_table(wb, ws1, ws2, ws2_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields)

    # Setup and call pivot_table
    ws2_name = 'NbrCommods'
    wb.Sheets.Add().Name = ws2_name
    ws2 = wb.Sheets(ws2_name)
    
    pt_name = 'NbrCommods'  # must be a string
    pt_rows = ['PERIOD']  # must be a list
    pt_cols = ['UNIT1']  # must be a list
    pt_filters = ['CTY_RPT', 'CH']  # must be a list
    # [0]: field name [1]: pivot table column name [3]: calulation method [4]: number format
    pt_fields = [  # must be a list of lists
                 ['NbrCommods', 'Sum of NbrCommods', win32c.xlSum, '#,###'],
                ]
    pivot_table(wb, ws1, ws2, ws2_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields)

    # Setup and call pivot_table
    ws2_name = 'QTY2'
    wb.Sheets.Add().Name = ws2_name
    ws2 = wb.Sheets(ws2_name)
    
    pt_name = 'QTY2'  # must be a string
    pt_rows = ['PERIOD']  # must be a list
    pt_cols = ['UNIT2']  # must be a list
    pt_filters = ['CTY_RPT', 'CH']  # must be a list
    # [0]: field name [1]: pivot table column name [3]: calulation method [4]: number format
    pt_fields = [  # must be a list of lists
                 ['QTY2', 'Sum of QTY2', win32c.xlSum, '#,###.00'],
                ]
    pivot_table(wb, ws1, ws2, ws2_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields)
    
    # Setup and call pivot_table
    ws2_name = 'QTY1'
    wb.Sheets.Add().Name = ws2_name
    ws2 = wb.Sheets(ws2_name)
    
    pt_name = 'QTY1'  # must be a string
    pt_rows = ['PERIOD']  # must be a list
    pt_cols = ['UNIT1']  # must be a list
    pt_filters = ['CTY_RPT', 'CH']  # must be a list
    # [0]: field name [1]: pivot table column name [3]: calulation method [4]: number format
    pt_fields = [  # must be a list of lists
                 ['QTY1', 'Sum of QTY1', win32c.xlSum, '#,###.00'],
                ]
    pivot_table(wb, ws1, ws2, ws2_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields)
    
    wb.Close(True)
    excel.Quit()



def makeForm(root, fields=fields):
    global savePathInput
    entries = {}
    # connects to DB3 grabbing TDM data availability
    
    for field in fields: 
        row = Frame(root)
        lbl = Label(row, width=22, text=field+': ', anchor='w') 

        if field == 'Save Path': 
            browseBTN = Button(row, text="Browse", command=browse_button)
            browseBTN.pack(side = RIGHT, padx=5, pady=5)

            savePathInput = Entry(row)
            savePathInput.insert(0, basePath)
            savePathInput.pack(side=RIGHT, expand=YES, fill=X)
            entries[field] = savePathInput
        elif field == 'SQL Code': 
            SQLcode = Text(row, width=5, height=10)
            SQLcode.pack(side=RIGHT, expand=YES, fill=X)
            entries[field] = SQLcode
        elif field == 'Server': 
            default = StringVar(root)
            current = default.get()
            dbCombobox = ttk.Combobox(row, textvariable=current)
            dbCombobox['values'] = ('SEVENFARMS_DB1', 
            # 'SEVENFARMS_DB2', 
            'SEVENFARMS_DB3')
            dbCombobox['state'] = 'readonly'
            dbCombobox.current(0)
            dbCombobox.pack(side=RIGHT, expand=YES, fill=X)
            entries[field] = dbCombobox
        else: 
            input = Entry(row)
            # input.insert(0, field)
            input.pack(side=RIGHT, expand=YES, fill=X)
            entries[field] = input

        row.pack(side=TOP, fill=X, expand=False, padx=5, pady=10)
        lbl.pack(side=LEFT)
        


    return entries

def browse_button():
    # Allow user to select a directory and store it in global var
    # called folder_path
    global folder_path
    filename = filedialog.askdirectory()
    # print(filename)
    savePathInput.insert(0, filename)

def validate_var(f_path, f_name, sheet_name, sqlQuery): 
    error = False
    # print(f_path)
    # print(f_name)
    # print(sheet_name)
    # print(sqlQuery)
    if len(str(f_path)) == 0 or len(str(f_path)) == 1: 
        error = 'Save path required.'
    if not os.path.exists(f_path): 
        error = 'Given file path does not exist or is not a folder.'
    if len(f_name.replace('.xlsx', '')) == 0: 
        error = 'File name required.'
    if len(sheet_name) == 0: 
        error = 'Sheet name required.'
    if len(sqlQuery) == 0:
        error = 'SQL statement required.'
    return error

def generateUnitChecks(root, e):
    # print(e['Excel File Name'].get())
    # print(e['Excel Sheet Name'].get())
    # print(e['Save Path'].get())
    # print(e['Server'].get())
    # print(e['SQL Code'].get('1.0', 'end-1c'))
    f_name = f"{e['Excel File Name'].get()}.xlsx"
    sheet_name = e['Excel Sheet Name'].get()
    f_path = Path(rf"{e['Save Path'].get()}")
    server = e['Server'].get()
    sqlQuery = e['SQL Code'].get('1.0', 'end-1c')

    error = validate_var(f_path, f_name, sheet_name, sqlQuery)
    if error == False: 
        pass
    else: 
        messagebox.showerror('Field Validation Error', error)
        return 
    create_excel_file(f_path, f_name, sheet_name, server, sqlQuery)
    run_excel(f_path, f_name, sheet_name)

if __name__ == '__main__':
    root = Tk()
    root.title('Unit Check Generator')
    root.geometry("1000x450")
    root.resizable(False, False)
    entries = makeForm(root, fields)
    root.bind('<Return>', (lambda event, e = entries: fetch(e)))

    
    
    row = Frame(root)
    exitBTN = Button(row, text = 'Quit', command = root.quit)
    exitBTN.pack(side = RIGHT, padx = 5, pady = 5)
    generateBTN = Button(row, text = 'Generate File', command = (lambda e = entries: generateUnitChecks(root, e)))
    generateBTN.pack(side = RIGHT, padx = 5, pady = 5)
    row.pack(side=BOTTOM, fill=X, expand=False, padx=5, pady=10)
    
     
    root.mainloop()
    