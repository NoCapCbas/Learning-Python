import pandas as pd
import os




path_to_create_txt = r"C:\\Users\\DDiaz\\Documents"

#type in folder name
folder_name = 'Dropbox'

#iterates through files if file name ends with .xlsx
directory_x = f'C:\\Users\\DDiaz\\Documents\\{folder_name}'
for file in os.listdir(directory_x):
    if file.endswith(".xlsx"):
        print(f'--{file[:-5]}')

        path_to_excel = f"{directory_x}\\{file}"


        read_file = pd.read_excel(path_to_excel)
        read_file.to_csv(f"{directory_x}\\{file[:-5]}.csv", index=0)
