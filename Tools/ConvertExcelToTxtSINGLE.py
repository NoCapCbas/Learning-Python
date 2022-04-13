from openpyxl import load_workbook
from openpyxl import Workbook
import pandas as pd

#Variable path example format

path_to_excel = r"C:\Users\DDiaz\Documents\Dropbox"
path_to_create_txt = r"C:\Users\DDiaz\Documents\Dropbox"

#path_to_excel used to read_excel and print to txt
xls = pd.ExcelFile(path_to_excel)
sheet_list = xls.sheet_names

for x in range(1,len(sheet_list)+1):



    #names .txt file
    txt_file_name = f'\\{sheet_list[x-1]}.txt'

    df = pd.read_excel(xls, sheet_list[x-1], )
    with open(path_to_create_txt + txt_file_name, 'w', encoding='utf8') as outfile:
        df.to_string(outfile)
