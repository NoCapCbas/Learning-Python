import win32com.client as win32
from pywintypes import com_error
from pathlib import Path
import sys
import pandas as pd  # only used for synthetic data
import numpy as np  # only used for synthetic data
import random  # only used for synthetic data
from datetime import datetime  # only used for synthetic data
# https://trenton3983.github.io/files/solutions/2020-06-22_pivot_table_win32com/create_pivot_table_with_win32com.html
# sql server import
import pyodbc
win32c = win32.constants
def create_excel_file(f_path: Path, f_name: str, sheet_name: str):
    
    filename = f_path / f_name
    
    # connects to DB3 grabbing TDM data availability
    conn = pyodbc.connect(
                            'Driver={SQL Server};'
                            'Server=SEVENFARMS_DB3;'
                            'Database=Control;'
                            'UID=sa;'
                            'PWD=Harpua88;'
                            'Trusted_Connection=No;'
    )
    
    df = pd.read_sql_query(f"""
    SELECT PERIOD, COUNT(DISTINCT commodity) AS NbrCommods,unit2 as UNIT2,sum(qty2) as QTY2,sum(value) as USD,cty_rpt as CTY_RPT
, SUM(QTY1) AS QTY1, UNIT1,SUBSTRING(COMMODITY,1,2) AS CH
  FROM [SRC_UN].[dbo].[E8] 
where cty_rpt in ('KG')
GROUP BY PERIOD,unit1,SUBSTRING(COMMODITY,1,2) ,unit2,cty_rpt
ORDER BY unit1,period""", conn)

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

def main():
    # sheet name for data
    sheet_name = 'I8'  # update with sheet name from your file
    # file path
    # f_path = Path.cwd()  # file in current working directory
    f_path = Path(r'C:\Users\DDiaz.ANH\Documents')  # file located somewhere else
    # excel file
    f_name = 'I8_UZ.xlsx'
    
    # create_excel_file(f_path, f_name, sheet_name)
    run_excel(f_path, f_name, sheet_name)
main()