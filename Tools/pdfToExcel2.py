import tabula
from bs4 import BeautifulSoup as bs
import pandas as pd
import html
import numpy as np
# importing all the required modules
import PyPDF2


path = r"C:\Users\DDiaz\Documents\Dropbox\Peru\arancel-ad-2022.pdf"
savePath = r"C:\Users\DDiaz\Documents\Dropbox\Peru\arancel-ad-2022.csv"

MASTER = pd.DataFrame(columns=['COMMODITY',
                                'DESCRIPTION',
                                'AV'
                            ])


# print('Converting...')
# tabula.convert_into(path, savePath, output_format="csv",lattice = False, guess=False, pages='all')

df = pd.read_csv(savePath, encoding = 'latin-1', sep='|', header=None, names=['Column'])
iList = []
# print(df)

two = '1'
for index, row in df.iterrows():
    if index < 475:
        continue

    # tools
    rowStr = row['Column']
    rowList = rowStr.split(',')
    # print(row['Column'])

    # if rowStr[:2].isnumeric() == True and df['Column'][index + 1]:
    #

    # if rowStr[:2].isnumeric() == True or '-' in rowList:
    #     # print(int(rowStr[:2]))
    #     if rowStr[:2].isnumeric() == True:
    #         qualifier = int(rowStr[:2])
    # else:
    #     continue

    if rowStr[:2].isnumeric() == True and rowStr[3:5].isnumeric():
        # commodity grab
        commodity =  str(rowStr[:2] + '.' + rowStr[3:5])
        # end commodity grab

        # av grab
        av = ''
        # end av grab

        # desc grab
        desc = rowStr[6:].replace(',', ' ')
        # end desc grab

    elif rowStr[:4].isnumeric() == True and rowStr[5:7].isnumeric() == True:
        # commodity grab
        for chrI in range(len(rowStr)-1):
            # print(rowStr[chrI])
            if rowStr[chrI].isnumeric() == True or rowStr[chrI] == '.':
                pass
            else:
                stopPoint = chrI
                break
        commodity = rowStr[:stopPoint]
        # end commodity grab

        # av grab
        if rowList[-1].isnumeric() == True:
            av = rowList[-1]
            # desc grab
            desc = rowStr[8:-1].replace(',', ' ')
            if rowList[-2].isnumeric() == True:
                av = rowList[-2]
                desc = rowStr[8:-3].replace(',', ' ')
        elif rowList[-1] == '':
            av = ''
            # desc grab
            desc = rowStr[8:].replace(',', ' ')
        else:

            av = ''
            # desc grab
            desc = rowStr[8:].replace(',', ' ')
        # end av grab

    elif '-' in rowStr:
        # commodity grab
        commodity = ''
        # end commodity grab

        # av grab
        if rowList[-1] != '':
            for chrI in reversed(range(len(rowList[-1])-1)):
                # print(rowStr[chrI])
                if rowStr[chrI].isnumeric() == True:
                    pass
                else:
                    stopavPoint = chrI
                    break
            av = rowStr[stopavPoint:]
        else:
            av = ''
        # end av grab

        # desc grab
        desc = rowStr.replace(',', '')
        # end desc grab
    else:
        continue

    # last desc check
    if desc[0].isnumeric() == True:
        condition = False
        for chrI in range(len(desc)-1):
            # print(rowStr[chrI])
            if desc[chrI].isnumeric() == True or desc[chrI] == '.':
                pass
            else:
                stopPoint = chrI
                break
        desc = desc[stopPoint:]
    count = 0
    for char in desc:
        if char.isnumeric():
            count +=1
        if count > 50:
            continue
    # last av check
    if len(desc) > 1:
        if av == '' and desc[-2].isnumeric() == True:
            av = desc[-2]
            desc = desc[:-3]
    if av.isnumeric() == False and av != '':
        continue
    if len(av) > 4:
        continue
    # last commodity check

    # print(two)
    if two == '':
        pass
    elif two.isnumeric() == True and commodity[:2].isnumeric() == True :
        if int(two) < int(commodity[:2]) or int(two) > int(commodity[:2]) + 2:
            continue

    intCount = 0
    count = 0
    for char in commodity:
        count += 1
        if char.isnumeric() == True:
            intCount += 1
        if count > 13:
            continue
        if intCount >= 10:
            continue
    # if two != '':
    two = commodity[:2]



    print(f'index: {index}')
    print(row['Column'])
    print(f'commodity: {commodity}')
    print(f'desc: {desc}')
    print(f'av: {av}')
    MASTER = MASTER.append({'COMMODITY': commodity, 'DESCRIPTION': desc, 'AV': av}, ignore_index=True)



    # if index == 800:
    #     break
    print()
print('Dumping to csv...')
# dfConverted.to_csv(f"C:\\Users\\DDiaz\\Documents\\Dropbox\\2020_12_imp_cty.csv", encoding='utf-8-sig')
MASTER.to_excel(f"C:\\Users\\DDiaz\\Documents\\Dropbox\\PERU.xlsx")
print('Script Executed Successfully.')
