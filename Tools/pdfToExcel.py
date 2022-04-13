import tabula
from bs4 import BeautifulSoup as bs
import pandas as pd
import html
import numpy as np


# path = "Y:\\Vietnam\\Vietnam_preliminary_PDF_dec_2021\\2021_12_imp_cty.pdf"
# path = "Y:\\Vietnam\\Vietnam_preliminary_PDF_dec_2021\\2021_12_exp_cty.pdf"
path = "Y:\\Vietnam\\Vietnam_preliminary_PDF_dec_2021\\2020_12_imp_cty.pdf"
# path = "Y:\\Vietnam\\Vietnam_preliminary_PDF_dec_2021\\2020_12_exp_cty.pdf"
savePath = "C:\\Users\\DDiaz\\Documents\\Dropbox\\test.csv"

df = tabula.read_pdf(path, pages = 'all')[1]
# print(len(df))
# print(df)
# convert PDF into CSV
tabula.convert_into(path, savePath, output_format="csv", stream=True,pages='all')
header_list = ['Unnamed: 0', 'Unnamed: 1', 'Reporting month', 'Year to date']
# read csv
dataCSV = pd.read_csv(savePath, delimiter=',', names=header_list)
# print(dataCSV)

dfConverted = pd.DataFrame(columns = ['Country', 'Main Imports', 'Units', 'Volume', 'Value(USD)', 'Volume2', 'Value(USD)2'])
dfConverted['Volume'] = dfConverted['Volume'].astype(str)
dfConverted['Value(USD)'] = dfConverted['Value(USD)'].astype(str)
dfConverted['Volume2'] = dfConverted['Volume2'].astype(str)
dfConverted['Value(USD)2'] = dfConverted['Value(USD)2'].astype(str)

for index, row in dataCSV.iterrows():

    if int(index) > 1:
        print(index)
        if str(row['Unnamed: 1']) == 'nan' and str(row['Reporting month']) == 'nan' and str(row['Year to date']) == 'nan':
            mainImportsAttch = row['Unnamed: 0']
            # print(mainImportsAttch)
            # print(dataCSV.iloc[[index-1]]['Unnamed: 0'])
            dataCSV.at[index-1,'Unnamed: 0'] = dataCSV.iloc[[int(index)-1]]['Unnamed: 0'] + mainImportsAttch
            # print(row['Unnamed: 0'])
            # print('Skipped')
            continue
        if row['Unnamed: 0'] == 'Country/Territory-Main imports':
            continue
        if row['Unnamed: 0'] == 'nan':
            continue


        # print(f"\trow 1: {row['Unnamed: 0']}")
        mainImports = row['Unnamed: 0']
        print(f'\tmainImports: {mainImports}')

        # print(f"\trow 2: {row['Unnamed: 1']}")
        units = row['Unnamed: 1']
        print(f'\tunits: {units}')

        # print(f"\trow 3: {row['Reporting month']}")
        cellSplit1 = str(row['Reporting month']).split(' ')
        # print(cellSplit1)
        if not cellSplit1 or cellSplit1[0] == 'nan':
            volume = '0'
            value = '0'
        if len(cellSplit1) == 1 and cellSplit1[0] != 'nan':
            value = cellSplit1[0]
            volume = '0'
        if len(cellSplit1) > 1:
            volume = cellSplit1[0]
            value = cellSplit1[1]
        print(f'\tvol: {volume}')
        print(f'\tval: {value}')

        # print(f"\trow 4: {row['Year to date']}")
        cellSplit2 = str(row['Year to date']).split(' ')
        # print(cellSplit2)
        if not cellSplit2 or cellSplit2[0] == 'nan':
            volume2 = '0'
            value2 = '0'
        if len(cellSplit2) == 1 and cellSplit2[0] != 'nan':
            value2 = cellSplit2[0]
            volume2 = '0'
        if len(cellSplit2) > 1:
            volume2 = cellSplit2[0]
            value2 = cellSplit2[1]
        print(f'\tvol2: {volume2}')
        print(f'\tval2: {value2}')
        if str(row['Unnamed: 1']) == 'nan' and volume == '0' and value != '0' and volume2 == '0' and value2 != '0':
            country = row["Unnamed: 0"]
            print(f'Country: {row["Unnamed: 0"]}')
        if volume == 'Reporting' or volume == 'Volume':
            continue
        print()


        dfConverted = dfConverted.append({'Country':country, 'Main Imports': mainImports , 'Units': units, 'Volume':volume, 'Value(USD)':value, 'Volume2':volume2, 'Value(USD)2':value2}, ignore_index=True)
        print('\n')
print(dfConverted)
print('Dumping to csv...')
# dfConverted.to_csv(f"C:\\Users\\DDiaz\\Documents\\Dropbox\\2020_12_imp_cty.csv", encoding='utf-8-sig')
dfConverted.to_excel(f"C:\\Users\\DDiaz\\Documents\\Dropbox\\2020_12_imp_cty.xlsx")
print('Script Executed Successfully.')
