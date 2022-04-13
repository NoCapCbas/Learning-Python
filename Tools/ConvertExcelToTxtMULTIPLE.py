from openpyxl import load_workbook
from openpyxl import Workbook
import pandas as pd
import os

path_to_create_txt = r"C:\\Users\\DDiaz\\Documents"

#type in folder name
folder_name = 'Dropbox'

#iterates through files if file name ends with .xlsx
directory_x = f'C:\\Users\\DDiaz\\Documents\\{folder_name}'
for file in os.listdir(directory_x):
    if file.endswith(".xlsx"):
        print(f'--{file}')

        path_to_excel = f"{directory_x}\\{file}"


        #path_to_excel used to read_excel and print to txt
        xls = pd.ExcelFile(path_to_excel)
        sheet_list = xls.sheet_names

        for x in range(1,len(sheet_list)+1):



            #names .txt file
            txt_file_name = f'\\{sheet_list[x-1]}.txt'

            df = pd.read_excel(xls, sheet_list[x-1], )
            with open(path_to_create_txt + txt_file_name, 'w', encoding='utf8') as outfile:
                df.to_string(outfile)
