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
# compares current DB3 availability tables of sourceName UN Comtrade to UN Comtrade Data availability API
apiKey = 'ToUiiGf8sA9Hozu0bQPD3S7fDeo1tAreJeMUY0QuNjxsTNe21yUEa22runfRBA4HP2X0MqQdjTovBx58f3nRubCooAAEjXsMMIakQS51l56K/eX6ypM7/F3KZmRa9WFi4tuWiZ0vCXXIUNdF/3MEtM9hBTgQ998ycs2mEBcHN/w='

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
# file1 = DataFile('Grenada', 308, 202110, True, True, 'New Download' )
# print(f'{file1.name} {file1.ctyCode} {file1.period} {file1.getExport()} {file1.getImport()} {file1.status}')


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
    # dfFileSorted = pd.DataFrame()
    port = 465
    smtp_server = "smtp.gmail.com"
    sender_email = "tdmUsageAlert@gmail.com"
    recipients = ["DDiaz@tradedatamonitor.com",
    # "y.zeng@tradedatamonitor.com",
    # "a.chan@tradedatamonitor.com",
    # "j.smith@tradedatamonitor.com"
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
            dfFile = dfFile.append({
            'COUNTRY': f'{dataFileClassList[k].ctyName} {dataFileClassList[k].ctyCode}',
            'PERIOD' : dataFileClassList[k].period,
            'AVAILABLE FLOWS' : flowText,
            'STATUS' : dataFileClassList[k].status[:len(dataFileClassList[k].status)-1]}, ignore_index=True)

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
    print('UNcomtradeAutomation Complete.')






if __name__ == '__main__':
    UNcomtradeAvailability()
