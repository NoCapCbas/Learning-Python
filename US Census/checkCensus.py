import os
import csv
import pandas as pd
monthCONC = {
'January':1,
'February':2,
'March':3,
'April':4,
'May':5,
'June':6,
'July':7,
'August':8,
'September':9,
'October':10,
'November':11,
'December':12,
}
stateCONC = {
                        
				"Alabama":'1',
							
				"Alaska":'2',
							
				"Arizona":'3',
							
				"Arkansas":'4', 
							
				"California":'5',
							
				"Colorado":'6', 
							
				"Connecticut":'7',
							
				"Delaware":'8',
							
				"Dist of Columbia":'9',
							
				"Florida":'10',
							
				"Georgia":'11', 
							
				"Hawaii":'12', 
							
				"Idaho":'13', 
							
				"Illinois":'14',
							
				"Indiana":'15',
							
				"Iowa":'16',
							
				"Kansas":'17',
							
				"Kentucky":'18',
							
				"Louisiana":'19',
							
				"Maine":'20',
							
				"Maryland":'21',
							
				"Massachusetts":'22',
							
				"Michigan":'23',
							
				"Minnesota":'24',
							
				"Mississippi":'25',
							
				"Missouri":'26',
							
				"Montana":'27',
							
				"Nebraska":'28',
							
				"Nevada":'29',
							
				"New Hampshire":'30',
							
				"New Jersey":'31', 
							
				"New Mexico":'32',
							
				"New York":'33', 
							
				"North Carolina":'34',
							
				"North Dakota":'35',
							
				"Ohio":'36',
							
				"Oklahoma":'37',
							
				"Oregon":'38',
							
				"Pennsylvania":'39',
							
				"Puerto Rico":'40',
							
				"Rhode Island":'41', 
							
				"South Carolina":'42',
							
				"South Dakota":'43',
							
				"Tennessee":'44',
							
				"Texas":'45',
							
				"US Virgin Islands":'46',
							
				"Utah":'47',
							
				"Vermont":'48',
							
				"Virginia":'49',
							
				"Washington":'50',
							
				"West Virginia":'51',
							
				"Wisconsin":'52',

                "Wyoming":'53',

                "Unknown":'54'
}
df = pd.DataFrame(columns=['FILE STATE', 'FILE CH', 'FILE TIME', 'STATE', 'CH', 'TIME', 'PATH'])
path = 'Y:\\_PyScripts\\Damon\\Log\\usCensus\\CensusEXPORTS'
yearPathList = os.listdir(path)
downloaded = []
downloadMistakes = []
for year in yearPathList: 
    monthPathList = os.listdir(os.path.join(path, year))
    for month in monthPathList:
        statePathList = os.listdir(os.path.join(f'{path}\\{year}', month))
        for state in statePathList: 
            fileList = os.listdir(os.path.join(f'{path}\\{year}\\{month}', state))
            for file in fileList:
                fileName = file.split('\\')[-1][:-4]
                fileYR = fileName[-4:]
                fileState = fileName.split('_')[1]
                fileMonth = fileName.split('_')[2].split(' ')[0]
                fileCH = fileName[:2]
                codeMonth = monthCONC[fileMonth]
                codeCH = int(fileCH)
                if codeCH > 77:
                    codeCH = codeCH - 1

                CODE = f'{int(fileYR[-1])-1}.{codeMonth}.{stateCONC[fileState]}.{codeCH}'
                downloaded.append(CODE)

                csvFile = open(os.path.join(f'{path}\\{year}\\{month}\\{state}', file))
                csvreader = csv.reader(csvFile)
                rowCount = 0
                for row in csvreader:
                    rowCount += 1 
                    if rowCount <= 3:
                        continue
                    elif rowCount > 4: 
                        break
                    else:
                        stateName = row[0]
                        ch = row[1][:2]
                        time = row[3]

                        

                        if stateName != fileState or ch != fileCH or time != f'{fileMonth} {fileYR}':
                            print(f'{path}\\{year}\\{month}\\{state}\\{file}')
                            print(f'{stateName} : {fileState}')
                            print(f'{ch} : {fileCH}')
                            print(f'{time} : {fileMonth} {fileYR}')
                            downloadMistakes.append(CODE)
                            
                            print()
                            df = pd.concat([df, pd.DataFrame({'FILE STATE': [fileState], 'FILE CH': [fileCH], 'FILE TIME': [f'{fileMonth} {fileYR}'], 'STATE': [stateName], 'CH':[ch], 'TIME':[time], 'PATH': [f'{path}\\{year}\\{month}\\{state}\\{file}'], 'DOWNLOAD CODE': [CODE]})], axis=0,ignore_index=True)
                            df.to_csv('C:\\Users\\DDiaz.ANH\\Documents\\checkCensus.csv', encoding='utf-8-sig', index=False)


# prints  codes with mistakes
# print(df['DOWNLOAD CODE'].to_list())

 

# read codes and write them all back except the ones that meet if condition
# lines = []
# file = "C:\\Users\\DDiaz.ANH\\Documents\\downloadLog.txt"
# with open(file, 'r') as f: 
#     lines = f.readlines()


# with open(file, 'w') as f: 
#     for num, line in enumerate(lines): 

#         if line[0] != '1':
#             f.write(line) 