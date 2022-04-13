# sql server import
import pyodbc
# pandas import
import pandas as pd
# datetime import
from datetime import datetime
from dateutil.relativedelta import relativedelta
# email import
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
# misc import
import shutil
import glob
import os
from six.moves import urllib
import requests
from time import sleep
import json
import time
import win32com.client as win32
from pywintypes import com_error
from pathlib import Path
import sys
import numpy as np

win32c = win32.constants
# compares current DB3 availability tables of sourceName UN Comtrade to UN Comtrade Data availability API
apiKey = 'ToUiiGf8sA9Hozu0bQPD3S7fDeo1tAreJeMUY0QuNjxsTNe21yUEa22runfRBA4HP2X0MqQdjTovBx58f3nRubCooAAEjXsMMIakQS51l56K/eX6ypM7/F3KZmRa9WFi4tuWiZ0vCXXIUNdF/3MEtM9hBTgQ998ycs2mEBcHN/w='

rootPath = 'Y:\\UN Comtrade\\ScrapedData\\'

class DataFile:
    def __init__(self, ctyName, ctyCode, period, exportFlow, importFlow, status):
        self.ctyName = ctyName
        self.ctyCode = ctyCode
        self.period = period
        self.exportFlow = exportFlow
        self.importFlow = importFlow
        self.status = status

    def getExport(self):
        if self.exportFlow == True:
            return 'Export'
        else:
            return ''

    def getImport(self):
        if self.importFlow == True:
            return 'Import'
        else:
            return ''

    def getFrequency(self):
        if len(self.period) == 4:
            return 'ANNUAL'
        elif len(self.period) == 6:
            return 'MONTHLY'
# file1 = DataFile('Grenada', 308, 202110, True, True, 'New Download' )
# print(f'{file1.name} {file1.ctyCode} {file1.period} {file1.getExport()} {file1.getImport()} {file1.status}')
def StatusEmail(phase, text):
	port = 465
	smtp_server = "smtp.gmail.com"
	sender_email = "tdmUsageAlert@gmail.com"
	recipients = ["DDiaz@tradedatamonitor.com",
				# "tdmdata@tradedatamonitor.com"
				# ,"m.alomar@tradedatamonitor.com"
				# ,"y.zeng@tradedatamonitor.com"
				# ,"a.chan@tradedatamonitor.com"
				# ,"j.smith@tradedatamonitor.com"
				]
	password = "tdm12345$"

	message = MIMEMultipart("alternative")
	message["Subject"] = f'UN Comtrade {phase}'
	message["from"] = "TDM Data Team"
	message["To"] = ", ".join(recipients)

	html = f"""\
    <html>
        <body>
            <p>
                {text}
            </p>


        </body>
      </html>
      """
	part1 = MIMEText(html, "html")

	message.attach(part1)

	context = ssl.create_default_context()

	with smtplib.SMTP_SSL(smtp_server,port, context = context) as server:
		server.login("tdmUsageAlert@gmail.com", password)
		server.sendmail(sender_email, recipients, message.as_string())

def connectDB3(TDMModule):
    # connects to DB3 grabbing TDM data availability
    conn = pyodbc.connect(
                            'Driver={SQL Server};'
                            'Server=SEVENFARMS_DB3;'
                            'Database=Control;'
                            'UID=sa;'
                            'PWD=Harpua88;'
                            'Trusted_Connection=No;'
    )

    dfTDM = pd.read_sql_query(f"""
    SELECT DISTINCT
    c.[CTY_DESC] AS [CTY_RPT],
    b.[DA_ISO_CODE2] AS [CTY_ISO],
    MIN(a.[StartYM]) AS [StartYM],
    MAX(a.[StopYM]) AS [StopYM],
    a.[DA_ISO_CODE3_NUMERIC] AS [CTY_CODE]
    FROM [Control].[dbo].[{TDMModule}] a
    LEFT JOIN [Control].[dbo].[{TDMModule}] b
    ON a.[DA_ISO_CODE3_NUMERIC] = b.[DA_ISO_CODE3_NUMERIC] AND a.[StartYM] = b.[StartYM] AND a.[StopYM] = b.[StopYM]
    LEFT JOIN [SEVENFARMS_DB1].[SP_MASTER].[dbo].[CTY_MASTER] c
    ON b.[DA_ISO_CODE2] = c.[CTY_ISO]
    WHERE a.[DA_ISO_CODE3_NUMERIC] != '' AND a.[DA_ISO_CODE3_NUMERIC] IS NOT NULL AND a.[SourceName] = 'UN Comtrade'
    GROUP BY a.[DA_ISO_CODE3_NUMERIC], b.[DA_ISO_CODE2], c.[CTY_DESC]""", conn)
    # removes rows with no country code
    dfTDM = dfTDM[dfTDM.CTY_CODE.notnull()]
    return dfTDM

def sendEmailOfDownloads(countriesDownloaded, failedDownload, dataFileClassList):
    print('Sending Email...')
    port = 465
    smtp_server = "smtp.gmail.com"
    sender_email = "tdmUsageAlert@gmail.com"
    recipients = ["DDiaz@tradedatamonitor.com",
    "y.zeng@tradedatamonitor.com",
    "a.chan@tradedatamonitor.com",
    "j.smith@tradedatamonitor.com"
    ]
    password = "tdm12345$"     #tdm12345$

    message = MIMEMultipart("alternative")
    message["Subject"] = 'UN Comtrade Report'
    message["from"] = "TDM Data Team"
    message["To"] = ", ".join(recipients)
    if dataFileClassList:
        str2 = ''
        dfFile = pd.DataFrame(columns=['COUNTRY', 'PERIOD', 'AVAILABLE FLOWS', 'STATUS'])
        for k in dataFileClassList:

            flowText = ''
            # print(dataFileClassList[k].exportFlow)
            if dataFileClassList[k].exportFlow and dataFileClassList[k].importFlow:
                flowText = 'Import|Export'
            elif dataFileClassList[k].importFlow:
                flowText = f'Import'
            elif dataFileClassList[k].exportFlow:
                flowText = f'Export'
            dfFile = pd.concat([dfFile, pd.DataFrame({
            'COUNTRY': [f'{dataFileClassList[k].ctyName} {dataFileClassList[k].ctyCode}'],
            'PERIOD' : [dataFileClassList[k].period],
            'AVAILABLE FLOWS' : [flowText],
            'STATUS' : [dataFileClassList[k].status[:len(dataFileClassList[k].status)-1]]})], ignore_index=True)

            str2 = f'{str2} {dataFileClassList[k].ctyName} {dataFileClassList[k].ctyCode} {dataFileClassList[k].period}&nbsp;&nbsp;&nbsp;{flowText}&nbsp;&nbsp;&nbsp;{dataFileClassList[k].status}<br>'
            # print(str2)
        # print(dfFile)
        dfFileSorted = dfFile.sort_values(['COUNTRY', 'PERIOD'], ascending=(True, True))
    if countriesDownloaded:
        str = ''
        for i in range(0,len(countriesDownloaded)):

            str = f'{str}{countriesDownloaded[i]}<br>'
            # print(countriesDownloaded[i][:-5])
    else:
        str = 'None'


    if failedDownload:
        strFailed = ''
        for i in range(0,len(failedDownload)):
            strFailed = f'{strFailed}{failedDownload[i]}<br>'
    else:
        strFailed = 'None'


    html = f"""\
    <html>
        <body>
            <p>
                {dfFileSorted.to_html(index=False)}
                <br>
                <hr>

            </p>

            <p>Failed Download:</p>
            <p>
                {strFailed}
            </p>

        </body>
      </html>
      """
    # print(html)
    content = MIMEText(html, "html")

    message.attach(content)

    context = ssl.create_default_context()

    with smtplib.SMTP_SSL("smtp.gmail.com",port, context = context) as server:
        server.login("tdmUsageAlert@gmail.com", password)
        server.sendmail(sender_email, recipients, message.as_string())

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


def getUNcomtradeDataAvailability(freq):
    print('-- Retrieving UN Comtrade Date Availability...')
    # connects to UN comtrade api grabbing data availability
    URL = f'http://comtrade.un.org/api//refs/da/view?type=C&freq={freq}&px=HS&token={apiKey}'


    response = requests.get(URL)

    # print(response.status_code)
    if response.status_code == 200:
        print('--Connection Success')
        sleep(1)

        print('----Grabbing UN Comtrade Data Availability...')

        text = response.text
        json_data = json.loads(text)
        #pprint(json_data)

        dfUNcomtrade = pd.DataFrame(json_data)
        # print(response.text)
        # print(response.content)
    else:
        dfUNcomtrade = None

    return dfUNcomtrade

def downloadCountries(MasterDownloadDic, freq, dfTDM, dfUN):
    print('Downloading New Countries Available...')
    countriesDownloaded = []
    failedDownload = []
    dataFileClassList = {}
    rg = [
            #{ "id": "all", "text": "All" },
            { "id": "1", "text": "Import" },
            { "id": "2", "text": "Export" },
            # { "id": "3", "text": "re-Export" },
            # { "id": "4", "text": "re-Import" }
        ]

    for ctyCode in MasterDownloadDic:
        downloadPeriodList = MasterDownloadDic.get(ctyCode)
        dfUN_Of_ctyCode = dfUN[dfUN.r == str(int(ctyCode))]
        # Grab current country Name
        ctyName = dfUN_Of_ctyCode.iloc[0]['rDesc']
        #print(downloadPeriodList)
        for period in downloadPeriodList:
            importExists = False
            exportExists = False
            dataFileStatus = ''

            for flow in range(0, len(rg)): #Flow Loop

                if rg[flow]["text"] == 'Import':
                    importExists = True
                    # print('importExists')
                if rg[flow]["text"] == 'Export':
                    exportExists = True
                    # print('exportExists')
                URL = f'http://comtrade.un.org/api/get?max=250000&type=C&freq={freq}&px=HS&ps={period}&r={ctyCode}&rg={rg[flow]["id"]}&cc=ag6&fmt=csv&token={apiKey}'
                if freq == 'M':
                    savePath = f'Y:\\UN Comtrade\\ScrapedData\\UNcomtradeMONTHLY\\{ctyName}\\{period}\\{ctyName}_{rg[flow]["text"]}_{period}.csv'
                    basePath = f'Y:\\UN Comtrade\\ScrapedData\\UNcomtradeMONTHLY\\{ctyName}\\{period}\\'
                else:  # freq == 'A':
                    savePath = f'Y:\\UN Comtrade\\ScrapedData\\UNcomtradeANNUAL\\{ctyName}\\{period}\\{ctyName}_{rg[flow]["text"]}_{period}.csv'
                    basePath = f'Y:\\UN Comtrade\\ScrapedData\\UNcomtradeANNUAL\\{ctyName}\\{period}\\'
                # confirming file does not already exist
                if(os.path.exists(savePath) and round(os.stat(savePath).st_size/1024) != 1):
                    # if savePath exists
                    countriesDownloaded.append(f'{ctyName} {ctyCode} {period} {rg[flow]["text"]}')
                    continue

                # confirming base path exists
                if(not os.path.exists(basePath)):
                    os.makedirs(basePath)

                ###########################################################################################################

                sleep(5)
                try:

                    #f = open(savePath, 'w')
                    print(f'Downloading...{ctyName}')
                    urllib.request.urlretrieve(URL, savePath)


                except Exception as e:

                    print(e)
                    print(f'----------TRY FAILED (Country Loop:{ctyName} {period} {rg[flow]["text"]})')

                list_of_files = glob.glob(basePath + '*.csv')

                if savePath in list_of_files:
                    if round(os.stat(savePath).st_size/1024) != 1:
                        # if new download is not 1 kb new data is available
                        countriesDownloaded.append(f'{ctyName} {ctyCode} {period} {rg[flow]["text"]} (New Download)')
                        dataFileStatus = f'{dataFileStatus}New {rg[flow]["text"]}|'
                    else:
                        # if new download is still 1kb no new data is available
                        countriesDownloaded.append(f'{ctyName} {ctyCode} {period} {rg[flow]["text"]}')
                    print(f'----Download Completed for {ctyName} {period} {freq} {rg[flow]["text"]}')

                else:
                    print(f'----Download failed for {ctyName} {period} {freq} {rg[flow]["text"]}')
                    failedDownload.append(f'{ctyName} {ctyCode} {period} {rg[flow]["text"]} (Failed Download)')
                    dataFileStatus = f'{dataFileStatus}Failed {rg[flow]["text"]}|'

                # verifies path exists or file is not 1kb if so there is no data for this flow
                if os.path.exists(savePath) == False or round(os.stat(savePath).st_size/1024) == 1:
                    if rg[flow]["text"] == 'Import':
                        importExists = False
                        # print('importExists')
                    if rg[flow]["text"] == 'Export':
                        exportExists = False



                ###########################################################################################################

            dataFileClassList[ctyName + period] = DataFile(ctyName, ctyCode, period, exportExists, importExists, dataFileStatus)


    return countriesDownloaded, failedDownload, dataFileClassList





def UNcomtradeAvailability():
    # main
    DataAvailabilityTables = ['Data_Availability_Monthly', 'Data_Availability_Annual']
    countriesDownloaded = []
    failedDownload = []
    dataFileClassList = {}


    for table in DataAvailabilityTables:
        print(f'UNcomtradeAvailability: Beginning {table}')
        MasterDownloadDic = {}
        # grabbing TDM dataframe
        dfTDM = connectDB3(table)


        # sets freq type to scrape un comtrade data availability
        if table == 'Data_Availability_Monthly':
            freq = 'M'

        if table == 'Data_Availability_Annual':
            freq = 'A'
        #print(freq)
        # grabbing UN dataframe
        dfUN = getUNcomtradeDataAvailability(freq)
        #print(dfUN)
        try:
            if dfUN == None:
                # Handles request error
                countriesDownloaded.append('***UN Comtrade Availability Request Error.***')
                continue
        except:
            pass

        for index, row in dfTDM.iterrows():


            ctyCode = row['CTY_CODE']
            ctyName = row['CTY_RPT']
            TDM_StartYM = row['StartYM']
            TDM_StopYM = row['StopYM']
            print('Processing ' + ctyName)

            dfUN_Of_ctyCode = dfUN[dfUN.r == str(int(ctyCode))]
            if dfUN_Of_ctyCode.empty:
                continue
            #print(dfUN_Of_ctyCode)


            # create UNcomtrade period List
            UNperiods = dfUN_Of_ctyCode['ps'].to_list()

            if freq == 'M':
                #create TDM period List
                TDMperiods = []
                cur_date = start = datetime.strptime(TDM_StartYM, '%Y%m').date()
                end = datetime.strptime(TDM_StopYM, '%Y%m').date()

                while cur_date <= end:
                    TDMperiods.append(str(cur_date)[:4] + str(cur_date)[5:-3])
                    cur_date += relativedelta(months=1)
            else:
                #create TDM period List
                TDMperiods = []
                cur_date = start = datetime.strptime(TDM_StartYM, '%Y').date()
                end = datetime.strptime(TDM_StopYM, '%Y').date()

                while cur_date <= end:
                    TDMperiods.append(str(cur_date)[:4])
                    cur_date += relativedelta(months=1)

            # compare period lists, subtract UN comtrade from TDM list to get UN comtrade periods that do not exist in TDM
            diffDownloadList = list(set(UNperiods) - set(TDMperiods))
            # if list is empty continue else grab new periods
            if diffDownloadList == []:
                continue

            else:
                # only grab periods that are new, from diffDownloadList grab periods above current TDM_StopYM
                newDataDownloadList = []
                for p in diffDownloadList:
                    if int(p) > int(TDM_StopYM):
                        newDataDownloadList.append(p)
                # if list is empty continue else assign list to ctyCode
                if newDataDownloadList == []:
                    continue
                MasterDownloadDic[ctyCode] = newDataDownloadList

        #print(MasterDownloadDic)
        newCtyList = downloadCountries(MasterDownloadDic, freq, dfTDM, dfUN)
        countriesDownloaded.extend(newCtyList[0])
        failedDownload.extend(newCtyList[1])
        dataFileClassList.update(newCtyList[2])


    sendEmailOfDownloads(countriesDownloaded, failedDownload, dataFileClassList)
    dataToLoad = []
    # Filter out new data files
    for file in dataFileClassList:
        # print(file)
        # print(f'status: {dataFileList[file].status}')
        # print(f'import: {dataFileList[file].getImport()}')
        # print(f'export: {dataFileList[file].getExport()}')
        # print(f'{rootPath}UNcomtrade{dataFileList[file].getFrequency()}\\{dataFileList[file].ctyName}\\{dataFileList[file].period}\\{dataFileList[file].ctyName}__{dataFileList[file].period}')
        # print()

        # check status of dataFile
        if dataFileClassList[file].status == '' or dataFileClassList[file].getFrequency == 'ANNUAL':
            continue
        else:
            # check Flows available
            if dataFileClassList[file].importFlow == True:
                tempPathToDataI = f'{rootPath}UNcomtrade{dataFileClassList[file].getFrequency()}\\{dataFileClassList[file].ctyName}\\{dataFileClassList[file].period}\\{dataFileClassList[file].ctyName}_{dataFileClassList[file].getImport()}_{dataFileClassList[file].period}.csv'
                if os.path.exists(tempPathToDataI):
                    dataToLoad.append(tempPathToDataI)
                
            if dataFileClassList[file].exportFlow == True:
                tempPathToDataE = f'{rootPath}UNcomtrade{dataFileClassList[file].getFrequency()}\\{dataFileClassList[file].ctyName}\\{dataFileClassList[file].period}\\{dataFileClassList[file].ctyName}_{dataFileClassList[file].getExport()}_{dataFileClassList[file].period}.csv'
                if os.path.exists(tempPathToDataE):
                    dataToLoad.append(tempPathToDataE)
    return dataToLoad


def getSql(sqlPath):
    # Open the external sql file.
    file = open(sqlPath, 'r')
    # Read out the sql script text in the file.
    sql = file.read()
    # Close the sql file object.
    file.close()
    return sql

def loadData(dataToLoad):
    print('Loading Data...')
    # Open SQL Server Connection
    conn = pyodbc.connect(
                            'Driver={SQL Server};'
                            'Server=SEVENFARMS_DB1;'
                            'Database=SRC_UN;'
                            'UID=sa;'
                            'PWD=Harpua88;'
                            'Trusted_Connection=No;', autocommit=True
    )
    cursor = conn.cursor()
    # Truncate UNcomtradeScraped
    cursor.execute('''
		TRUNCATE TABLE [SRC_UN].[dbo].[UNcomtradeScraped]
               ''')
    

    # Insert new data into truncated table
    UNcolumns = ['CLASSIFICATION'
      ,'YEAR'
      ,'PERIOD'
      ,'PERIOD_DESC'
      ,'AGGREGATE_LEVEL'
      ,'IS_LEAF_CODE'
      ,'TRADE_FLOW_CODE'
      ,'TRADE_FLOW'
      ,'REPORTER_CODE'
      ,'REPORTER'
      ,'REPORTER_ISO'
      ,'PARTNER_CODE'
      ,'PARTNER'
      ,'PARTNER_ISO'
      ,'SND_PARTNER_CODE'
      ,'SND_PARTNER'
      ,'SND_PARTNER_ISO'
      ,'CUSTOMS_PROC_CODE'
      ,'CUSTOMS'
      ,'MODE_OF_TRANSPORT_CODE'
      ,'MODE_OF_TRANSPORT'
      ,'COMMODITY_CODE'
      ,'COMMODITY'
      ,'QTY_UNIT_CODE'
      ,'QTY_UNIT'
      ,'QTY'
      ,'ALT_QTY_UNIT_CODE'
      ,'ALT_QTY_UNIT'
      ,'ALT_QTY'
      ,'NETWEIGHT'
      ,'GROSSWEIGHT'
      ,'TRADE_VALUE'
      ,'CIF_TRADE_VALUE'
      ,'FOB_TRADE_VALUE'
      ,'FLAG']
    # connects to DB3 grabbing TDM data availability

    for path in dataToLoad:
        data = pd.read_csv(path)
        df = data.astype(str)
        df.columns = UNcolumns
        df = df.fillna('')
        df = df.replace(np.nan, '')
        df = df.replace('nan', '')
        
        # Insert DataFrame to Table
        for row in df.itertuples():

            cursor.execute(f'''
                        INSERT INTO [SRC_UN].[dbo].[UNcomtradeScraped] ([FILENAME],[CLASSIFICATION]
                            ,[YEAR]
                            ,[PERIOD]
                            ,[PERIOD DESC]
                            ,[AGGREGATE LEVEL]
                            ,[IS LEAF CODE]
                            ,[TRADE FLOW CODE]
                            ,[TRADE FLOW]
                            ,[REPORTER CODE]
                            ,[REPORTER]
                            ,[REPORTER ISO]
                            ,[PARTNER CODE]
                            ,[PARTNER]
                            ,[PARTNER ISO]
                            ,[2ND PARTNER CODE]
                            ,[2ND PARTNER]
                            ,[2ND PARTNER ISO]
                            ,[CUSTOMS PROC. CODE]
                            ,[CUSTOMS]
                            ,[MODE OF TRANSPORT CODE]
                            ,[MODE OF TRANSPORT]
                            ,[COMMODITY CODE]
                            ,[COMMODITY]
                            ,[QTY UNIT CODE]
                            ,[QTY UNIT]
                            ,[QTY]
                            ,[ALT QTY UNIT CODE]
                            ,[ALT QTY UNIT]
                            ,[ALT QTY]
                            ,[NETWEIGHT(KG)]
                            ,[GROSSWEIGHT(KG)]
                            ,[TRADE VALUE(US$)]
                            ,[CIF TRADE VALUE(US$)]
                            ,[FOB TRADE VALUE(US$)]
                            ,[FLAG])
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', 
                        path,
                        row.CLASSIFICATION
                        ,row.YEAR
                        ,row.PERIOD
                        ,row.PERIOD_DESC
                        ,row.AGGREGATE_LEVEL
                        ,row.IS_LEAF_CODE
                        ,row.TRADE_FLOW_CODE
                        ,row.TRADE_FLOW
                        ,row.REPORTER_CODE
                        ,row.REPORTER
                        ,row.REPORTER_ISO
                        ,row.PARTNER_CODE
                        ,row.PARTNER
                        ,row.PARTNER_ISO
                        ,row.SND_PARTNER_CODE
                        ,row.SND_PARTNER
                        ,row.SND_PARTNER_ISO
                        ,row.CUSTOMS_PROC_CODE
                        ,row.CUSTOMS
                        ,row.MODE_OF_TRANSPORT_CODE
                        ,row.MODE_OF_TRANSPORT
                        ,row.COMMODITY_CODE
                        ,row.COMMODITY
                        ,row.QTY_UNIT_CODE
                        ,row.QTY_UNIT
                        ,row.QTY
                        ,row.ALT_QTY_UNIT_CODE
                        ,row.ALT_QTY_UNIT
                        ,row.ALT_QTY
                        ,row.NETWEIGHT
                        ,row.GROSSWEIGHT
                        ,row.TRADE_VALUE
                        ,row.CIF_TRADE_VALUE
                        ,row.FOB_TRADE_VALUE
                        ,row.FLAG
                        )
        
    return conn, cursor
    

def processData(conn, cursor):
    print('Processing Data...')
    sqlPath = "C:\\Users\\DDiaz.ANH\\Documents\\sqlScripts\\UNcomtrade\\MONTHLY\\MONTHLY_pyscript.sql"
    # Grabs sql script
    sql = getSql(sqlPath)
    # Execute the read out sql script string.
    cursor.execute(sql)
    return conn, cursor

def verifyData(conn, cursor):
    print('Verifying Data...')
    # Verify data does not already exist in E8 and I8
    print('\tChecking if data already Exists in E8 and I8...')
    query = pd.read_sql_query('''
    SELECT DISTINCT 
    A.CTY_RPT AS [CTY_RPT_STEP4], 
    A.[PERIOD] AS [PERIOD_STEP4],  
    B.CTY_RPT AS [CTY_RPT_E8],  
    B.[PERIOD] AS [PERIOD_E8],
    CASE
        WHEN (B.[CTY_RPT] IS NULL OR B.[PERIOD] IS  NULL)
        THEN 'DOES NOT EXIST IN E8' 
    ELSE 'EXISTS IN E8'
    END AS [CONDITION]
    FROM [SRC_UN].[dbo].[UNcomtradeMONTHLY-STEP4_EXP] A
    LEFT JOIN [SRC_UN].[dbo].[E8] B
    ON A.CTY_RPT = B.CTY_RPT AND A.[PERIOD] = B.[PERIOD]
    ORDER BY 1, 2
    ''', conn)

    VERIFYe8 = pd.DataFrame(query, columns=['CTY_RPT_STEP4', 'PERIOD_STEP4', 'CTY_RPT_E8', 'PERIOD_E8', 'CONDITION'])

    query = pd.read_sql_query('''
    SELECT DISTINCT 
    A.CTY_RPT AS [CTY_RPT_STEP4], 
    A.[PERIOD] AS [PERIOD_STEP4],  
    B.CTY_RPT AS [CTY_RPT_I8],  
    B.[PERIOD] AS [PERIOD_I8],
    CASE
        WHEN (B.[CTY_RPT] IS NULL OR B.[PERIOD] IS  NULL)
        THEN 'DOES NOT EXIST IN I8' 
    ELSE 'EXISTS IN I8'
    END AS [CONDITION]
    FROM [SRC_UN].[dbo].[UNcomtradeMONTHLY-STEP4_IMP] A
    LEFT JOIN [SRC_UN].[dbo].[I8] B
    ON A.CTY_RPT = B.CTY_RPT AND A.[PERIOD] = B.[PERIOD]
    ORDER BY 1, 2
    ''', conn)

    VERIFYi8 = pd.DataFrame(query, columns=['CTY_RPT_STEP4', 'PERIOD_STEP4', 'CTY_RPT_I8', 'PERIOD_I8', 'CONDITION'])

    # Generic checks
    print('\tGeneric Checks...')
    dfE8 = pd.read_sql_query('''
    SELECT * 
    FROM [SRC_UN].[dbo].[UNcomtradeMONTHLY-STEP4_EXP]
    ''', conn)
    dfI8 = pd.read_sql_query('''
    SELECT * 
    FROM [SRC_UN].[dbo].[UNcomtradeMONTHLY-STEP4_IMP]
    ''', conn)
    failedChecks = {'E8NULL_VALUES':[False], 
                    'I8NULL_VALUES':[False], 
                    'E8COMMODITY_LEN 8 DIGITS':[False],
                    'I8COMMODITY_LEN 8 DIGITS':[False] 
    }

    if dfE8.isnull().values.any() == True:
        failedChecks['E8NULL_VALUES'][0] = True
    if dfI8.isnull().values.any() == True:
        failedChecks['I8NULL_VALUES'][0] = True
    if dfE8['COMMODITY'].astype(str).map(len).nunique() != 1 and dfE8['COMMODITY'].astype(str).map(len).unique()[0] != 8:
        failedChecks['E8COMMODITY_LEN 8 DIGITS'][0] = True
    if dfI8['COMMODITY'].astype(str).map(len).nunique() != 1 and dfI8['COMMODITY'].astype(str).map(len).unique()[0] != 8:
        failedChecks['I8COMMODITY_LEN 8 DIGITS'][0] = True
    # print(failedChecks)
    genericChecks = pd.DataFrame.from_dict(failedChecks)
    # print(genericChecks)
    # cty_rptE8 = dfE8['CTY_RPT'].astype(str).unique()
    # cty_rptI8 = dfI8['CTY_RPT'].astype(str).unique()

    # Compare totals Before Processing and After Processing
    dfExportTotalsCheck = pd.read_sql_query('''
    SELECT B.[CTY_RPT], 
			B.[PERIOD], 
			[TRADE FLOW] AS [FLOW],
			[TOTAL VALUE] AS [TOTAL VALUE BEFORE], 
			SUM([VALUE]) AS [TOTAL VALUE AFTER],
			[TOTAL VALUE] - SUM([VALUE]) AS [DIFF],
			(([TOTAL VALUE] - SUM([VALUE]))/[TOTAL VALUE])*100 AS [PERCENT DIFF]
		FROM [SRC_UN].[dbo].[UNcomtradeScrapedTOTAL] A  
		LEFT JOIN [SRC_UN].[dbo].[UNcomtradeMONTHLY-STEP4_EXP] B
		ON A.[CTY_RPT] = B.[CTY_RPT] AND A.[PERIOD] = B.[PERIOD]
		WHERE [TRADE FLOW] = 'Exports'
		GROUP BY B.[CTY_RPT], B.[PERIOD], [trade flow], [TOTAL VALUE]
    ''', conn)
    dfImportTotalsCheck = pd.read_sql_query('''
    SELECT B.[CTY_RPT], 
			B.[PERIOD], 
			[TRADE FLOW] AS [FLOW],
			[TOTAL VALUE] AS [TOTAL VALUE BEFORE], 
			SUM([VALUE]) AS [TOTAL VALUE AFTER],
			[TOTAL VALUE] - SUM([VALUE]) AS [DIFF],
			(([TOTAL VALUE] - SUM([VALUE]))/[TOTAL VALUE])*100 AS [PERCENT DIFF]
		FROM [SRC_UN].[dbo].[UNcomtradeScrapedTOTAL] A  
		LEFT JOIN [SRC_UN].[dbo].[UNcomtradeMONTHLY-STEP4_IMP] B
		ON A.[CTY_RPT] = B.[CTY_RPT] AND A.[PERIOD] = B.[PERIOD]
		WHERE [TRADE FLOW] = 'Imports'
		GROUP BY B.[CTY_RPT], B.[PERIOD], [trade flow], [TOTAL VALUE]
    ''', conn)

    port = 465
    smtp_server = "smtp.gmail.com"
    sender_email = "tdmUsageAlert@gmail.com"
    recipients = ["DDiaz@tradedatamonitor.com",
                # "tdmdata@tradedatamonitor.com"
                # ,"m.alomar@tradedatamonitor.com"
                # ,"y.zeng@tradedatamonitor.com"
                # ,"a.chan@tradedatamonitor.com"
                # ,"j.smith@tradedatamonitor.com"
                ]
    password = "tdm12345$"

    message = MIMEMultipart("alternative")
    message["Subject"] = f'UN Comtrade Check'
    message["from"] = "TDM Data Team"
    message["To"] = ", ".join(recipients)

    html = f"""\
    <html>
        <body>
            <h3>Generic Checks</h3>
                {genericChecks.to_html(index=False)}
            <br>
            <hr>
            <h3>Totals Check: Export</h3>
                {dfExportTotalsCheck.to_html(index=False)}
            <br>
            <hr>
            <h3>Totals Check: Import</h3>
                {dfImportTotalsCheck.to_html(index=False)}
            <br>
            <hr>
            <h3>VERIFY if Data in E8</h3>
                {VERIFYe8.to_html(index=False)}
            <br>
            <hr>
            <h3>VERIFY if Data in I8</h3>
                {VERIFYi8.to_html(index=False)}
        </body>
    </html>
    """
    part1 = MIMEText(html, "html")

    message.attach(part1)

    context = ssl.create_default_context()

    with smtplib.SMTP_SSL(smtp_server,port,context=context) as server:
        server.login("tdmUsageAlert@gmail.com", password)
        server.sendmail(sender_email, recipients, message.as_string())
    return conn, cursor

def autoUnitCheck(conn, cursor): 
    f_path = Path(r'Y:\UN Comtrade\Unit Checks\New_damon\MONTHLY')  # file located somewhere else
    allCTYS = pd.read_sql_query('''
    SELECT DISTINCT CTY_RPT
    FROM [SRC_UN].[dbo].[UNcomtradeMONTHLY-STEP4_IMP]
    ''', conn)

    

    # # excel file
    
    for cty in allCTYS['CTY_RPT']:
        e_name = f'E8_{cty}.xlsx'
        i_name = f'I8_{cty}.xlsx'
        pivotDataE8 = pd.read_sql_query(f'''
        SELECT PERIOD, COUNT(DISTINCT commodity) AS NbrCommods,UNIT2,sum(qty2) as QTY2,sum(value) as USD,CTY_RPT AS CTY_RPT
        , SUM(QTY1) AS QTY1, UNIT1,SUBSTRING(COMMODITY,1,2) AS CH
        FROM [SRC_UN].[dbo].[E8] 
        where cty_rpt in ('{cty}')
        GROUP BY PERIOD,unit1,SUBSTRING(COMMODITY,1,2) ,unit2,cty_rpt
        ORDER BY unit1,period
        ''', conn)
        pivotDataI8 = pd.read_sql_query(f'''
        SELECT PERIOD, COUNT(DISTINCT commodity) AS NbrCommods,UNIT2,sum(qty2) as QTY2,sum(value) as USD,CTY_RPT AS CTY_RPT
        , SUM(QTY1) AS QTY1, UNIT1,SUBSTRING(COMMODITY,1,2) AS CH
        FROM [SRC_UN].[dbo].[I8] 
        where cty_rpt in ('{cty}')
        GROUP BY PERIOD,unit1,SUBSTRING(COMMODITY,1,2) ,unit2,cty_rpt
        ORDER BY unit1,period
        ''', conn)

        pivotDataE8.to_excel(f_path/e_name, sheet_name='E8', index=False)
        pivotDataI8.to_excel(f_path/i_name, sheet_name='I8', index=False)
        generateUnitChecks(f_path, e_name, 'E8')
        generateUnitChecks(f_path, i_name, 'I8')

def generateUnitChecks(f_path: Path, f_name: str, sheet_name: str):
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


    


if __name__ == '__main__':
    # Data Aquisition
    try: 
        dataToLoad = UNcomtradeAvailability()
    except Exception as e:
        StatusEmail('UN Comtrade UNcomtradeAvailability() Error', e)
        quit() 

    

    # Data Loading
    if dataToLoad:
        StatusEmail('Automation Launched', 'UN Comtrade Automation Launched.')

        try: 
            conn, cursor = loadData(dataToLoad)
            StatusEmail('loadData', 'Data Loaded Successfully.')
            print('\tData Loaded.')
        except Exception as e:
            StatusEmail('UN Comtrade loadData() Error', e)
            quit()
    
        # Data Processing
        try:
            conn, cursor = processData(conn, cursor)
            StatusEmail('processData', 'Data Proccessed Successfully.')
            print('\tData processed.')
        except Exception as e:
            StatusEmail('processData() Error', e)
            quit() 

        # Data Verification Checks
        try: 
            conn, cursor = verifyData(conn, cursor)
            conn, cursor = autoUnitCheck(conn, cursor)
            cursor.close()
            conn.close()
        except Exception as e:

            quit()


