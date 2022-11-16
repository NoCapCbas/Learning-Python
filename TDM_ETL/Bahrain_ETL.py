import encodings
import requests 
import pyodbc 
import pandas as pd
import logging
import glob
from datetime import datetime
import shutil
import traceback
import sys
import os
import json
import random
from time import sleep
from bs4 import BeautifulSoup as bs
from pywintypes import com_error
from pathlib import Path
from pandasql import sqldf
import win32com.client as win32
# email import
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
# from selenium import webdriver
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
win32c = win32.constants


class ETL():
	def __init__(self, country):
		self.country = country 
		
		conn = pyodbc.connect(
							'Driver={SQL Server};'
							'Server=SF-TEST\SFTESTDB;'
							'Database=Control;'
							'UID=sa;'
							'PWD=Harpua88;'
							'Trusted_Connection=No;'
		)
		Q = f"""
		SELECT StopYM
		from [Control].[dbo].[Data_Availability_Monthly]
		WHERE Declarant = '{self.country}'"""
		# int(list(pd.read_sql_query(Q, conn)['StopYM'])[0][-2:])
		self.lastPeriodOfTDMdata = list(pd.read_sql_query(Q, conn)['StopYM'])[0]
		
		conn.close()
		if int(self.lastPeriodOfTDMdata[-2:]) == 12:
			self.yearToQuery = str(int(self.lastPeriodOfTDMdata[:-2]) + 1)
			self.nextPeriod = '01'
		else:
			self.yearToQuery = self.lastPeriodOfTDMdata[:-2]
			self.nextMon = str(int(self.lastPeriodOfTDMdata[-2:]) + 1)
			if len(self.nextMon) == 1:
				self.nextMon = '0' + self.nextMon
		# print(self.yearToQuery)
		self.nextPeriodOfTDMdata = self.yearToQuery + str(self.nextMon)
		#self.nextPeriodOfTDMdata = '202206' # hard set for testing
		
		self.current_year = datetime.now().year
		self.today = datetime.now().strftime('%#m/%#d/%Y')
		# self.today = '10/16/2022' # hard set for testing 
		# URLS
		self.data_source_url = 'https://www.data.gov.bh/en/ResourceCenter'
		# PATHS
		self.logPath = f"Y:\\_PyScripts\\Damon\\{self.country}\\Log"
		self.downloadPath = f"Y:\\_PyScripts\\Damon\\{self.country}\\Downloads"
		self.archivePath = f'Y:\\{self.country}\\Archive\\{self.nextPeriodOfTDMdata}'
		self.datafilesPath = f'Y:\\{self.country}\\Data Files'
		self.excelBugPath = 'C:\\Users\\DDiaz.ANH\\AppData\\Local\\Temp\\gen_py\\3.10\\00020813-0000-0000-C000-000000000046x0x1x9'
		if os.path.exists(self.excelBugPath):
			shutil.rmtree(self.excelBugPath)
		# print(self.archivePath)
		if not os.path.exists(self.logPath):
			os.makedirs(self.logPath)
		if not os.path.exists(self.downloadPath): 
			os.makedirs(self.downloadPath)
		if not os.path.exists(self.archivePath): 
			os.makedirs(self.archivePath)
		
		# create logger
		self.logger = logging.getLogger('Log')
		self.logger.setLevel(logging.INFO)
		# create file handler and set level
		handler = logging.FileHandler(filename=self.logPath + "\\RUNTIME.log", mode='w')
		handler.setLevel(logging.INFO)
		# create formatter
		format = logging.Formatter('%(asctime)s %(levelname)s:%(message)s', datefmt='%b-%d-%Y %H:%M:%S')
		# add formatter to handler
		handler.setFormatter(format)
		# add handler to logger
		self.logger.addHandler(handler)
		# adding console logging
		consoleHandler = logging.StreamHandler()
		self.logger.addHandler(consoleHandler)
		self.logger.info('Bahrain_ETL')
		self.logger.info(f'Last Period Published: {self.lastPeriodOfTDMdata}')
	
	def StatusEmail(self, phase, text, text2 = '', ALL = False):
		# Grab RUNTIME.log
		with open(self.logPath + "\\RUNTIME.log", mode='r') as fileObj:
			RUNTIMElog = fileObj.read()

		port = 465
		smtp_server = "smtp.gmail.com"
		sender_email = "tdmUsageAlert@gmail.com"
		if ALL:
			recipients = ["DDiaz@tradedatamonitor.com"
						,"y.zeng@tradedatamonitor.com"
						,"a.chan@tradedatamonitor.com"
						,"j.smith@tradedatamonitor.com"
						]
		else:
			recipients = ["DDiaz@tradedatamonitor.com"]

		password = "tdm12345$"
		password = 'drwvgldfexinkmyc'
		message = MIMEMultipart("alternative")
		message["Subject"] = f'{self.country} {phase}'
		message["from"] = "TDM Data Team"
		message["To"] = ", ".join(recipients)

		attachment = MIMEApplication(RUNTIMElog)
		attachment['Content-Disposition'] = 'attachment; filename="RUNTIME.log"'
		message.attach(attachment)

		html = f"""\
		<html>
			<body>
				<h3>{text}</h3>
				<p>
					{text2}
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
	
	def generateUnitChecks(self, f_path: Path, f_name: str, sheet_name: str):
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
				print(f'Failed to open spreadsheet. Invalid filename or location: {filename}')
			else:
				raise e
			sys.exit(1)

		excel.Range("A2").Select()
		excel.ActiveWindow.FreezePanes = True
		
		# set worksheet
		ws1 = wb.Sheets(sheet_name)

		# Setup and call pivot_table
		ws2_name = 'CH'
		wb.Sheets.Add().Name = ws2_name
		ws2 = wb.Sheets(ws2_name)
		
		pt_name = 'CH' # must be a string
		pt_rows = ['PERIOD'] # must be a list
		pt_cols = ['CH'] # must be a list
		pt_filters = ['CTY_RPT'] # must be a list
		# [0]: field name [1]: pivot table column name [3]: calulation method [4]: number format
		pt_fields = [ # must be a list of lists
					['NbrCommods', 'Sum of NbrCommods', win32c.xlSum, '#,###'],
					]
		self.pivot_table(wb, ws1, ws2, ws2_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields)
		excel.Range("A5").Select()
		excel.ActiveWindow.FreezePanes = True

		# Setup and call pivot_table
		ws2_name = 'VALUE'
		wb.Sheets.Add().Name = ws2_name
		ws2 = wb.Sheets(ws2_name)
		
		pt_name = 'VALUE' # must be a string
		pt_rows = ['PERIOD'] # must be a list
		pt_cols = ['UNIT1'] # must be a list
		pt_filters = ['CTY_RPT', 'CH']  # must be a list
		# [0]: field name [1]: pivot table column name [3]: calulation method [4]: number format
		pt_fields = [  # must be a list of lists
					['USD', 'Sum of USD', win32c.xlSum, '$#,##0.00'],
					]
		self.pivot_table(wb, ws1, ws2, ws2_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields)
		excel.Range("A6").Select()
		excel.ActiveWindow.FreezePanes = True

		# Setup and call pivot_table
		ws2_name = 'NbrCommods'
		wb.Sheets.Add().Name = ws2_name
		ws2 = wb.Sheets(ws2_name)
		
		pt_name = 'NbrCommods'  # must be a string
		pt_rows = ['PERIOD']  # must be a list
		pt_cols = ['UNIT1']  # must be a list
		pt_filters = ['CTY_RPT', 'CH'] # must be a list
		# [0]: field name [1]: pivot table column name [3]: calulation method [4]: number format
		pt_fields = [  # must be a list of lists
					['NbrCommods', 'Sum of NbrCommods', win32c.xlSum, '#,###'],
					]
		self.pivot_table(wb, ws1, ws2, ws2_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields)
		excel.Range("A6").Select()
		excel.ActiveWindow.FreezePanes = True

		# Setup and call pivot_table
		ws2_name = 'QTY2'
		wb.Sheets.Add().Name = ws2_name
		ws2 = wb.Sheets(ws2_name)
		
		pt_name = 'QTY2'  # must be a string
		pt_rows = ['PERIOD']  # must be a list
		pt_cols = ['UNIT2']  # must be a list
		pt_filters = ['CTY_RPT', 'CH'] # must be a list
		# [0]: field name [1]: pivot table column name [3]: calulation method [4]: number format
		pt_fields = [  # must be a list of lists
					['QTY2', 'Sum of QTY2', win32c.xlSum, '#,###'],
					]
		self.pivot_table(wb, ws1, ws2, ws2_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields)
		excel.Range("A6").Select()
		excel.ActiveWindow.FreezePanes = True

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
					['QTY1', 'Sum of QTY1', win32c.xlSum, '#,###'],
					]
		self.pivot_table(wb, ws1, ws2, ws2_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields)
		excel.Range("A6").Select()
		excel.ActiveWindow.FreezePanes = True

		wb.Close(True)
		excel.Quit()

	def pivot_table(self,wb: object, ws1: object, pt_ws: object, ws_name: str, pt_name: str, pt_rows: list, pt_cols: list, pt_filters: list, pt_fields: list):
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
	
	def autoUnitCheck(self, conn, cursor):
		self.logger.info('\t\tWorking on Unit Checks...')
		if not os.path.exists(rf'Y:\{self.country}\Unit Checks\{self.nextPeriodOfTDMdata}'):
			os.mkdir(rf'Y:\{self.country}\Unit Checks\{self.nextPeriodOfTDMdata}')

		f_path = Path(rf'Y:\{self.country}\Unit Checks\{self.nextPeriodOfTDMdata}')
		
		table_matches = {
			'E':'TEMP_STEP5_EXP',
			'I':'TEMP_STEP5_IMP',
		}
		for k,v in table_matches.items():
			
			if os.path.exists(self.excelBugPath):
				shutil.rmtree(self.excelBugPath)
			file_name = f'{k}8.xlsx'
			self.logger.info(f'\t\t\t{file_name}')

			# [SRC_Japan].[dbo].[Imports_NEW]
			pivotData = pd.read_sql_query(f'''
			SELECT PERIOD, COUNT(DISTINCT commodity) AS NbrCommods,UNIT2,sum(qty2) as QTY2,sum(value) as USD,CTY_RPT AS CTY_RPT
			, SUM(QTY1) AS QTY1, UNIT1,SUBSTRING(COMMODITY,1,2) AS CH
			FROM [SRC_{self.country}].[dbo].[{k}8]
			GROUP BY PERIOD,unit1,SUBSTRING(COMMODITY,1,2) ,unit2,cty_rpt
			ORDER BY period
			''', conn)
			
			pivotData.to_excel(f_path/file_name, sheet_name=f'{k}8', index=False)
			self.generateUnitChecks(f_path, file_name, f'{k}8')
			if os.path.exists(self.excelBugPath):
				shutil.rmtree(self.excelBugPath)
		
		self.logger.info('\t\t\tUnit Checks Generated.')
		
		return conn, cursor
	
	def getSql(self, sqlPath):
		# Open the external sql file.
		file = open(sqlPath, 'r')
		# Read out the sql script text in the file.
		sql = file.read()
		# Close the sql file object.
		file.close()
		return sql

	def extractData(self):
		self.logger.info('\tExtracting Data...')
		
		# example of download url
		# https://www.data.gov.bh/en/ResourceCenter/DownloadFile?id=4061
		url = "https://www.data.gov.bh/en/ResourceCenter/GetRCFilesData"
		downloadURLS = []
		payload = json.dumps({
			"node": 760,
			"all": False,
			"search": "",
			"expanded": "",
			"listView": False
		})
		headers = {
			'authority': 'www.data.gov.bh',
			'__c$v$t': 'GrD-BpjikGEUlY3Y5921Sl5-egqiT1fxJfHc_CabusF6GeRPTVf8kxqBfnY1pjhzwKqiyBchDVFS73YYKW35povCrZ41',
			'__r$v$t': 'KCbgeViqZO36eYPdG4ffBxSqr93aXe5SKktqYyn3DM1jCXgCs9QKtHrr9vKKL_GGBBmcPsbupJrptO41L-FNR5JEF701',
			'accept': '*/*',
			'accept-language': 'en-US,en;q=0.9',
			'content-type': 'application/json',
			'cookie': '_ga=GA1.3.315297658.1660325382; _gid=GA1.3.460126948.1660325382; __utmc=121403037; __utmz=121403037.1660325382.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none); __utma=121403037.315297658.1660325382.1660325382.1660328607.2; __utmb=121403037.1.10.1660328607',
			'origin': 'https://www.data.gov.bh',
			'referer': 'https://www.data.gov.bh/en/ResourceCenter?id=3935&d=1',
			'sec-ch-ua': '"Chromium";v="104", " Not A;Brand";v="99", "Google Chrome";v="104"',
			'sec-ch-ua-mobile': '?0',
			'sec-ch-ua-platform': '"Windows"',
			'sec-fetch-dest': 'empty',
			'sec-fetch-mode': 'cors',
			'sec-fetch-site': 'same-origin',
			'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36',
			'x-requested-with': 'XMLHttpRequest'
		}
		r = requests.request('POST', url, headers=headers, data=payload)
		# under Economic Statistics
		nodesForeignTrade = r.json()['Nodes']
		for node in nodesForeignTrade:
			# print(node)
			if node['Text'] == str(self.nextPeriodOfTDMdata)[:-2]:

				# print(node)
				payload = json.dumps({
					"node": node['Key'],
					"all": False,
					"search": "",
					"expanded": "",
					"listView": False
				})
				r = requests.request('POST', url, headers=headers, data=payload)
				nodesYear = r.json()['Nodes']
				# for flow in year folder
				for flow in nodesYear:
					# print(flow)
					if flow['Text'] == 'IMPORT':
						payload = json.dumps({
							"node": flow['Key'],
							"all": False,
							"search": "",
							"expanded": "",
							"listView": False
							})
						r = requests.request('POST', url, headers=headers, data=payload)
						nodesOptions = r.json()['Nodes']
						# find BY COUNTRY & COMMODITY
						for node2 in nodesOptions:
							
							# print(node2)
							if node2['Text'] == 'BY COUNTRY & COMMODITY':
								payload = json.dumps({
								"node": node2['Key'],
								"all": False,
								"search": "",
								"expanded": "",
								"listView": False
								})
								r = requests.request('POST', url, headers=headers, data=payload)
								nodesFiles = r.json()['Nodes']
								htmlDate = nodesFiles[-1]['Columns'][2]
								latestPostedDataDate = htmlDate.replace('<div class="column_posted">', '').replace('</div>', '')
									
								if self.today == latestPostedDataDate:
									self.logger.info(f'\t\t{flow["Text"]}: Todays date ({self.today}) matches latest posted date ({latestPostedDataDate})')
									tempDownload = f'https://www.data.gov.bh/en/ResourceCenter/DownloadFile?id={nodesFiles[-1]["Key"]}'
									downloadURLS.append(tempDownload)
								else:
									self.logger.info(f'\t\t{flow["Text"]}: Todays date ({self.today}) does not match latest posted date ({latestPostedDataDate})')
					
					if flow['Text'] == 'EXPORT':
						payload = json.dumps({
							"node": flow['Key'],
							"all": False,
							"search": "",
							"expanded": "",
							"listView": False
							})
						r = requests.request('POST', url, headers=headers, data=payload)
						nodesOptions = r.json()['Nodes']
						# for folder in export folder
						for node2 in nodesOptions:
							# print(node2)
							payload = json.dumps({
							"node": node2['Key'],
							"all": False,
							"search": "",
							"expanded": "",
							"listView": False
							})
							r = requests.request('POST', url, headers=headers, data=payload)
							nodesOptions2 = r.json()['Nodes']

							for node3 in nodesOptions2:
							
								# print(node3)
								if node3['Text'] == 'BY COUNTRY & COMMODITY':
									payload = json.dumps({
									"node": node3['Key'],
									"all": False,
									"search": "",
									"expanded": "",
									"listView": False
									})
									r = requests.request('POST', url, headers=headers, data=payload)
									nodesFiles = r.json()['Nodes']
									htmlDate = nodesFiles[-1]['Columns'][2]
									latestPostedDataDate = htmlDate.replace('<div class="column_posted">', '').replace('</div>', '')
										
									if self.today == latestPostedDataDate:
										self.logger.info(f'\t\t{flow["Text"]}: Todays date ({self.today}) matches latest posted date ({latestPostedDataDate})')
										tempDownload = f'https://www.data.gov.bh/en/ResourceCenter/DownloadFile?id={nodesFiles[-1]["Key"]}'
										downloadURLS.append(tempDownload)
									else:
										self.logger.info(f'\t\t{flow["Text"]}: Todays date ({self.today}) does not match latest posted date ({latestPostedDataDate})')
						

				for downloadable in downloadURLS:
					r = requests.request('POST', downloadable, headers=headers)
					# print(r.headers)
					filename = r.headers['Content-Disposition'].replace('attachment; filename="', '').replace('"', '')
					self.logger.info(f'\t\t\t{filename} Downloading...')
					open(self.downloadPath + f'\\{filename}', 'wb').write(r.content)

		if len(os.listdir(self.downloadPath)) == 4:
			self.logger.info('\t\t\t*New Data Available')
			self.StatusEmail('New Data', f'New {self.country} Data Downloading', '')
			dataToLoad = True
		elif len(os.listdir(self.downloadPath)) > 0:
			self.logger.info(f'\t\t\t{len(os.listdir(self.downloadPath))} files in Download Folder.')
			self.StatusEmail('Potential New Data', f'{self.country} {len(os.listdir(self.downloadPath))} files in Download Folder. \nLooking for {self.nextPeriodOfTDMdata} Data.', '')
			dataToLoad = False
		else:
			self.logger.info('\t\t\tNo New Data Available')
			dataToLoad = False

		self.logger.info('\tData Extraction Complete.')
		return dataToLoad

	def loadData(self):
		self.logger.info('\tLoading Data...')

		# Create Connection to [SRC_Bahrain]
		conn = pyodbc.connect(
			driver = '{ODBC Driver 17 for SQL Server}', 
			server = 'SF-PROCESS', 
			database = 'SRC_Bahrain',
			uid = 'sa', 
			pwd = 'Harpua88'
		)
		cursor = conn.cursor()

		MASTER = pd.DataFrame()

		for file in os.listdir(self.downloadPath): 
			self.logger.info(f'\t\tReading {file}...')
			
			tempXL = pd.read_excel(f'{self.downloadPath}\\{file}', keep_default_na=False)
			
			startI = tempXL.index[tempXL[tempXL.columns[0]] == 'Commodity No'].tolist()[0]
			# print(startI)
			tempXL = tempXL[startI+1:]
			# print(tempXL)
			# CSV file to archive 
			tempXL.to_csv(self.archivePath + f'\\{file.replace("xlsx", "txt")}', index=False)
			tempXL.insert(0, 'PERIOD', f'{self.nextPeriodOfTDMdata}')
			tempXL.insert(0, 'FILENAME', f'{self.archivePath}\\{file}')
			
			# Add temp to MASTER
			MASTER = pd.concat([MASTER, tempXL], ignore_index=True)
			# moves excel to archive
			shutil.move(f'{self.downloadPath}\\{file}', f'{self.archivePath}\\{file}')
		self.logger.info('\t\tAll Data Read.')
		# print(MASTER)
		MASTER = MASTER.astype(str)
		MASTER = MASTER.rename(columns={
			MASTER.columns[1]:'PERIOD', 
			MASTER.columns[2]:'COMMODITY', 
			MASTER.columns[3]:'DESC', 
			MASTER.columns[4]:'UN CODE', 
			MASTER.columns[5]:'CTY ABBR', 
			MASTER.columns[6]:'CTY DESC', 
			MASTER.columns[7]:'VALUE BD', 
			MASTER.columns[8]:'VALUE USD', 
			MASTER.columns[9]:'WEIGHT KG', 
			MASTER.columns[10]:'QTY', 
			MASTER.columns[11]:'UM', 
			MASTER.columns[12]:'ARABIC DESC', 
		})
		# print(MASTER)

		# Load MASTER to DB
		self.logger.info('\t\t\tDROPPING TABLE SRC_Bahrain.dbo.TEMP')
		cursor.execute(f"IF OBJECT_ID('SRC_Bahrain.dbo.TEMP') IS NOT NULL DROP TABLE [SRC_Bahrain].[dbo].[TEMP];")
		cursor.commit()

		# CREATE TABLE
		create_statements_cols = ''
		for col in MASTER.columns:
			if MASTER.columns[-1] == col:
				create_statements_cols = create_statements_cols + f'[{col}] [nvarchar](max) NULL'
			else: 
				create_statements_cols = create_statements_cols + f'[{col}] [nvarchar](max) NULL,'

		self.logger.info(f'\t\t\tCREATING TABLE SRC_Bahrain.dbo.TEMP')
		cursor.execute(f"""
		CREATE TABLE SRC_Bahrain.dbo.TEMP(
			{create_statements_cols}
		)
		""")
		cursor.commit()

		# Loop through list of split df and insert into table
		self.logger.info(f'\t\t\tINSERTING {len(MASTER)} ROWS INTO SRC_Bahrain.dbo.TEMP')
		self.logger.info('')
		insert_to_temp_table = f'INSERT INTO SRC_Bahrain.dbo.TEMP VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)'
		cursor.fast_executemany = True

		# INSERT INTO TEMP
		masterValues = MASTER.values.tolist()
		rowStepper = 1000000
		# Inserting 1000000 rows of data at a time to prevent error
		for rowI in range(0,len(masterValues), rowStepper):
			cursor.executemany(insert_to_temp_table, masterValues[rowI:rowI+rowStepper])
			cursor.commit()

		cursor.close()
		conn.close()
		self.logger.info('\tData Loaded.')

	def processData(self):
		self.logger.info('\tProcessing Data...')
		# Grab SQL processing script
		sqlScriptPath = f"Y:\\_PyScripts\\Damon\\{self.country}\\Bahrain_pyscript.sql"
		sql = self.getSql(sqlScriptPath)
		# Create Connection
		conn = pyodbc.connect(
			driver = '{ODBC Driver 17 for SQL Server}', 
			server = 'SF-PROCESS',
			database = 'SRC_Bahrain', 
			uid = 'sa', 
			pwd = 'Harpua88'
		)
		cursor = conn.cursor()

		# Execute the read out sql script string
		cursor.execute(sql)
		cursor.commit()
		running = True
		while running == True:
			sleep(3)
			try:
				query = cursor.execute('''
				SELECT *
				FROM [SRC_Bahrain].[DBO].[RUNNING-STATUS];
				''').fetchone()
			
			except:
				continue
			STATUS = query[0]
			if STATUS == 0:
				self.logger.info('\t\t\tSQL Script Executed.')
				running = False 
				cursor.execute('''
				IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[SRC_Bahrain].[dbo].[RUNNING-STATUS]') AND type in (N'U'))
				DROP TABLE [SRC_Bahrain].[DBO].[RUNNING-STATUS];
				''')
				cursor.commit()
			else:
				running = True

		cursor.close()
		conn.close()
		self.logger.info('\tData Processed.')

	def dataChecks(self):
		self.logger.info('\tPerforming Checks...')

		# Create Connection
		conn = pyodbc.connect(
			driver = '{ODBC Driver 17 for SQL Server}', 
			server = 'SF-PROCESS',
			database = 'SRC_Bahrain', 
			uid = 'sa', 
			pwd = 'Harpua88'
		)
		cursor = conn.cursor()

		# Generic Checks
		self.logger.info('\t\tGeneric Checks...')

		PRE_FINAL_TABLE_EXP = '[SRC_Bahrain].[dbo].[TEMP_STEP5_EXP]'
		PRE_FINAL_TABLE_IMP = '[SRC_Bahrain].[dbo].[TEMP_STEP5_IMP]'

		dfE8 = pd.read_sql_query(f'''
		SELECT * 
		FROM {PRE_FINAL_TABLE_EXP};
		''', conn)
		dfI8 = pd.read_sql_query(f'''
		SELECT * 
		FROM {PRE_FINAL_TABLE_IMP};
		''', conn)
		failedChecks = {
			'CHECK PERFORMED': ['STATUS'], 
			'E8 NULL VALUES DO NOT EXIST': [False], 
			'I8 NULL VALUES DO NOT EXIST': [False], 
			'E8 COMMODITY_LEN IS 8 DIGITS': [False], 
			'I8 COMMODITY_LEN IS 8 DIGITS': [False],
			'ALL E8 CTY_PTN EXIST IN CTY_MASTER DB1': [False],
			'ALL I8 CTY_PTN EXIST IN CTY_MASTER DB1': [False], 
			'ALL E8 UNIT1 EXIST IN UOM_MASTER DB1': [False], 
			'ALL I8 UNIT1 EXIST IN UOM_MASTER DB1': [False], 
			'ALL E8 UNIT2 EXIST IN UOM_MASTER DB1': [False], 
			'ALL I8 UNIT2 EXIST IN UOM_MASTER DB1': [False], 
		}

		if dfE8.isnull().values.any() != True:
			failedChecks['E8 NULL VALUES DO NOT EXIST'][0] = True

		if dfI8.isnull().values.any() != True:
			failedChecks['I8 NULL VALUES DO NOT EXIST'][0] = True

		if dfE8['COMMODITY'].astype(str).map(len).nunique() == 1 and dfE8['COMMODITY'].astype(str).map(len).unique()[0] == 8:
			failedChecks['E8 COMMODITY_LEN IS 8 DIGITS'][0] = True

		if dfI8['COMMODITY'].astype(str).map(len).nunique() == 1 and dfI8['COMMODITY'].astype(str).map(len).unique()[0] == 8:
			failedChecks['I8 COMMODITY_LEN IS 8 DIGITS'][0] = True

		# Check CTY_PTN 
		tempE8 = pd.read_sql_query(f'''
		SELECT DISTINCT CTY_PTN, CTY_ABBR
		FROM {PRE_FINAL_TABLE_EXP} A
		LEFT JOIN [SP_MASTER].[dbo].[CTY_MASTER] B
		ON A.CTY_PTN = B.CTY_ABBR
		''', conn)
		if tempE8.isnull().values.any() != True: 
			failedChecks['ALL E8 CTY_PTN EXIST IN CTY_MASTER DB1'][0] = True

		tempI8 = pd.read_sql_query(f'''
		SELECT DISTINCT CTY_PTN, CTY_ABBR
		FROM {PRE_FINAL_TABLE_IMP} A
		LEFT JOIN [SP_MASTER].[dbo].[CTY_MASTER] B
		ON A.CTY_PTN = B.CTY_ABBR
		''', conn)
		if tempI8.isnull().values.any() != True: 
			failedChecks['ALL I8 CTY_PTN EXIST IN CTY_MASTER DB1'][0] = True

		# Check UNIT1
		tempE8 = pd.read_sql_query(f'''
		SELECT DISTINCT UNIT1, ANH_UOM
		FROM {PRE_FINAL_TABLE_EXP} A
		LEFT JOIN [SP_MASTER].[dbo].[UOM_MASTER] B
		ON A.UNIT1 = B.ANH_UOM
		''', conn)
		if tempE8.isnull().values.any() != True: 
			failedChecks['ALL E8 UNIT1 EXIST IN UOM_MASTER DB1'][0] = True

		tempI8 = pd.read_sql_query(f'''
		SELECT DISTINCT UNIT1, ANH_UOM
		FROM {PRE_FINAL_TABLE_IMP} A
		LEFT JOIN [SP_MASTER].[dbo].[UOM_MASTER] B
		ON A.UNIT1 = B.ANH_UOM
		''', conn)
		if tempI8.isnull().values.any() != True: 
			failedChecks['ALL I8 UNIT1 EXIST IN UOM_MASTER DB1'][0] = True

		# Check UNIT2 
		tempE8 = pd.read_sql_query(f'''
		SELECT DISTINCT UNIT2, ANH_UOM
		FROM {PRE_FINAL_TABLE_EXP} A
		LEFT JOIN [SP_MASTER].[dbo].[UOM_MASTER] B
		ON A.UNIT2 = B.ANH_UOM
		''', conn)
		if tempE8.isnull().values.any() != True: 
			failedChecks['ALL E8 UNIT2 EXIST IN UOM_MASTER DB1'][0] = True

		tempI8 = pd.read_sql_query(f'''
		SELECT DISTINCT UNIT2, ANH_UOM
		FROM {PRE_FINAL_TABLE_IMP} A
		LEFT JOIN [SP_MASTER].[dbo].[UOM_MASTER] B
		ON A.UNIT2 = B.ANH_UOM
		''', conn)
		if tempI8.isnull().values.any() != True: 
			failedChecks['ALL I8 UNIT2 EXIST IN UOM_MASTER DB1'][0] = True

		# print(failedChecks)
		genericChecks = pd.DataFrame.from_dict(failedChecks)
		genericChecks = genericChecks.set_index('CHECK PERFORMED').transpose()
		# print(genericChecks)

		# TOTAL CHECKS
		self.logger.info('\t\tQuerying Totals...')
		EXP_TOTALS_SRC = pd.read_sql_query(f'''
		SELECT
			CAST([VALUE BD] AS VARCHAR) AS [SRC TOTAL_BD]
			,CAST([VALUE USD] AS VARCHAR) AS [SRC TOTAL]
			,CAST([QTY] AS VARCHAR) AS [SRC QTY1]
			,CAST([WEIGHT KG] AS VARCHAR) AS [SRC QTY2]
		FROM [SRC_Bahrain].[dbo].[TEMP]
		WHERE [COMMODITY] = '' AND [FILENAME] LIKE '%TOTAL EXPORT%'
		''',conn)
		EXP_TOTALS = pd.read_sql_query(f'''
		SELECT 
			CAST(SUM([VALUE_BD]) AS VARCHAR) AS [TOTAL_BD]
			,CAST(SUM([VALUE]) AS VARCHAR) AS [TOTAL]
			,CAST(SUM([QTY1]) AS VARCHAR) AS [QTY1]
			,CAST(SUM([QTY2]) AS VARCHAR) AS [QTY2]
		FROM {PRE_FINAL_TABLE_EXP}
		''',conn)
		IMP_TOTALS_SRC = pd.read_sql_query(f'''
		SELECT
			CAST([VALUE BD] AS VARCHAR) AS [SRC TOTAL_BD]
			,CAST([VALUE USD] AS VARCHAR) AS [SRC TOTAL]
			,CAST([QTY] AS VARCHAR) AS [SRC QTY1]
			,CAST([WEIGHT KG] AS VARCHAR) AS [SRC QTY2]
		FROM [SRC_Bahrain].[dbo].[TEMP]
		WHERE [COMMODITY] LIKE '%TOTAL%' AND [FILENAME] LIKE '%IMPORT%'
		''',conn)
		IMP_TOTALS = pd.read_sql_query(f'''
		SELECT 
			CAST(SUM([VALUE_BD]) AS VARCHAR) AS [TOTAL_BD]
			,CAST(SUM([VALUE]) AS VARCHAR) AS [TOTAL]
			,CAST(SUM([QTY1]) AS VARCHAR) AS [QTY1]
			,CAST(SUM([QTY2]) AS VARCHAR) AS [QTY2]
		FROM {PRE_FINAL_TABLE_IMP}
		''',conn)
		
		IMP_TOTALS = pd.concat([IMP_TOTALS_SRC, IMP_TOTALS], axis=1, ignore_index=True)
		IMP_TOTALS = IMP_TOTALS.rename(columns={
			IMP_TOTALS.columns[0]:'SRC TOTAL_BD',
			IMP_TOTALS.columns[1]:'SRC TOTAL',
			IMP_TOTALS.columns[2]:'SRC QTY1',
			IMP_TOTALS.columns[3]:'SRC QTY2',
			IMP_TOTALS.columns[4]:'PRC TOTAL_BD',
			IMP_TOTALS.columns[5]:'PRC TOTAL',
			IMP_TOTALS.columns[6]:'PRC QTY1',
			IMP_TOTALS.columns[7]:'PRC QTY2',
		})

		EXP_TOTALS = pd.concat([EXP_TOTALS_SRC, EXP_TOTALS], axis=1, ignore_index=True)
		EXP_TOTALS = IMP_TOTALS.rename(columns={
			EXP_TOTALS.columns[0]:'SRC TOTAL_BD',
			EXP_TOTALS.columns[1]:'SRC TOTAL',
			EXP_TOTALS.columns[2]:'SRC QTY1',
			EXP_TOTALS.columns[3]:'SRC QTY2',
			EXP_TOTALS.columns[4]:'PRC TOTAL_BD',
			EXP_TOTALS.columns[5]:'PRC TOTAL',
			EXP_TOTALS.columns[6]:'PRC QTY1',
			EXP_TOTALS.columns[7]:'PRC QTY2',
		})

		# Check if newest period exists within final tables before inserting, want a table of length zero 
		check_E8 = pd.read_sql_query(f'''
		SELECT DISTINCT a.PERIOD, b.PERIOD
		FROM {PRE_FINAL_TABLE_EXP} a
		LEFT JOIN [SRC_Bahrain].[dbo].[E8] b
		ON a.PERIOD = b.PERIOD
		WHERE b.PERIOD IS NOT NULL;
		''', conn)
		# print(len(check_E8))
		check_I8 = pd.read_sql_query(f'''
		SELECT DISTINCT a.PERIOD, b.PERIOD
		FROM {PRE_FINAL_TABLE_IMP} a
		LEFT JOIN [SRC_Bahrain].[dbo].[I8] b
		ON a.PERIOD = b.PERIOD
		WHERE b.PERIOD IS NOT NULL;
		''', conn)
		# print(len(check_I8))
		
		insrtCondition = False
		if list(genericChecks.all())[0] and len(check_E8) == 0 and len(check_I8) == 0:
			# INSERT INTO ARCHIVE 
			cursor.execute(f'''
			-- EXPORT
			INSERT INTO [SRC_Bahrain].[dbo].[SRC_EXP_IMP_RXP_ARCHIVE]
			SELECT cast([FILENAME] as VARCHAR(1000))
				,cast([flow] as varchar(3))
				,cast([PERIOD] as varchar(6))
				,cast([COMMODITY] as nvarchar(255))
				,cast([DESC] as nvarchar(255))
				,cast([UN CODE] as nvarchar(255))
				,[CTY ABBR]
				,[CTY DESC]
				,[VALUE BD]
				,[VALUE USD]
				,[WEIGHT KG]
				,[QTY]
				,[UM]
				,cast([ARABIC DESC] as nvarchar(255))
			FROM [SRC_Bahrain].[dbo].[TEMP_STEP0_EXP];

			-- IMPORT
			INSERT INTO [SRC_Bahrain].[dbo].[SRC_EXP_IMP_RXP_ARCHIVE]
			SELECT cast([FILENAME] as VARCHAR(1000))
				,cast([flow] as varchar(3))
				,cast([PERIOD] as varchar(6))
				,cast([COMMODITY] as nvarchar(255))
				,cast([DESC] as nvarchar(255))
				,cast([UN CODE] as nvarchar(255))
				,[CTY ABBR]
				,[CTY DESC]
				,[VALUE BD]
				,[VALUE USD]
				,[WEIGHT KG]
				,[QTY]
				,[UM]
				,cast([ARABIC DESC] as nvarchar(255))
			FROM [SRC_Bahrain].[dbo].[TEMP_STEP0_IMP];
			''')
			cursor.commit()

			# insert into final tables
			cursor.execute(f'''
			-- EXPORT
			INSERT INTO [SRC_Bahrain].[dbo].[E8]
			SELECT [CTY_RPT]
				[CTY_RPT]
				,[CTY_PTN]
				,[COMMODITY]
				,[PERIOD]
				,[YR]
				,[MON]
				,[VALUE_BD]
				,[VALUE]
				,[UNIT1]
				,[QTY1]
				,[UNIT2]
				,[QTY2]
			FROM {PRE_FINAL_TABLE_EXP};

			-- IMPORT
			INSERT INTO [SRC_Bahrain].[dbo].[I8]
			SELECT [CTY_RPT]
				[CTY_RPT]
				,[CTY_PTN]
				,[COMMODITY]
				,[PERIOD]
				,[YR]
				,[MON]
				,[VALUE_BD]
				,[VALUE]
				,[UNIT1]
				,[QTY1]
				,[UNIT2]
				,[QTY2]
			FROM {PRE_FINAL_TABLE_IMP};
			''')
			cursor.commit()

			# drop tables
			cursor.execute('''
			--DROP TABLE [SRC_Bahrain].[dbo].[TEMP];
			DROP TABLE [SRC_Bahrain].[dbo].[TEMP_STEP0_EXP];
			DROP TABLE [SRC_Bahrain].[dbo].[TEMP_STEP0_IMP];
			DROP TABLE [SRC_Bahrain].[dbo].[TEMP_STEP1_EXP];
			DROP TABLE [SRC_Bahrain].[dbo].[TEMP_STEP1_IMP];
			DROP TABLE [SRC_Bahrain].[dbo].[TEMP_STEP2_EXP];
			DROP TABLE [SRC_Bahrain].[dbo].[TEMP_STEP2_IMP];
			DROP TABLE [SRC_Bahrain].[dbo].[TEMP_STEP3_EXP];
			DROP TABLE [SRC_Bahrain].[dbo].[TEMP_STEP3_IMP];
			DROP TABLE [SRC_Bahrain].[dbo].[TEMP_STEP4_EXP];
			DROP TABLE [SRC_Bahrain].[dbo].[TEMP_STEP4_IMP];
			DROP TABLE [SRC_Bahrain].[dbo].[TEMP_STEP5_EXP];
			DROP TABLE [SRC_Bahrain].[dbo].[TEMP_STEP5_IMP];
			''')
			cursor.commit()

			# Unit Checks
			try:
				conn, cursor = self.autoUnitCheck(conn, cursor)
			except Exception as e:
				self.StatusEmail('autoUnitCheck() Error', e, traceback.format_exc())
			insrtCondition = True
			shutil.copy(f'{self.logPath}\\RUNTIME.log', f'{self.logPath}\\{self.nextPeriodOfTDMdata}.log')
			self.logger.info('Data Verified.')
		else:
			self.logger.info(f'INSERT STATEMENT EXECUTED: {insrtCondition}')
		
		# Send Check Email
		port = 465
		recipients = ["DDiaz@tradedatamonitor.com"
					,"y.zeng@tradedatamonitor.com"
					,"a.chan@tradedatamonitor.com"
					,"j.smith@tradedatamonitor.com"
					,"m.alomar@tradedatamonitor.com"
					]
		smtp_server = "smtp.gmail.com"
		sender_email = "tdmUsageAlert@gmail.com"
		password = "tdm12345$"
		password = 'drwvgldfexinkmyc'
		message = MIMEMultipart("alternative")
		message["Subject"] = f'{self.country} Check'
		message["from"] = "TDM Data Team"
		message["To"] = ", ".join(recipients)

		html = f"""\
		<html>
			<body>
				<table style="width:100%">
					<tbody>
						<tr>
							<td>{genericChecks.to_html().replace('<td>True', '<td><p class = "grn">True').replace('<td>False', '<td><p class = "rd">False').replace('</td>', '</p></td>')}</td>

						</tr>
					</tbody>
				</table>
				<br>

				<hr>
				<h3>Totals: EXPORT</h3>
					{EXP_TOTALS.to_html(index=False)}
				<br>
				<hr>
				<h3>Totals: IMPORT</h3>
					{IMP_TOTALS.to_html(index=False)}

				<br>
				<hr>
				<p>INSERT STATEMENT EXECUTED: {str(insrtCondition).replace('True', '<p class = "grn">True</p>').replace('False', '<p class = "rd">False</p>')}</p>
				
			</body>
			<style>
				p.grn {{
					color: green;
				}}
				p.rd {{
					color: red;
				}}
			</style>
		</html>
		"""
		# print(html)
		part1 = MIMEText(html, "html")
		message.attach(part1)
		context = ssl.create_default_context()
		with smtplib.SMTP_SSL(smtp_server,port,context=context) as server:
			server.login("tdmUsageAlert@gmail.com", password)
			server.sendmail(sender_email, recipients, message.as_string())

		if insrtCondition == False:
			ValueError(f'INSERT STATEMENT EXECUTED: {insrtCondition}')

		cursor.close()
		conn.close()

		self.logger.info('\tChecks Performed.')


	def run(self):
		self.logger.info(f'Launching {self.country} ETL...')
		
		# Data Extraction
		try:
			dataToLoad = self.extractData()
		except Exception as e:
			self.logger.exception('extractData() Error')
			self.StatusEmail('extractData() Error', e, traceback.format_exc())
			sys.exit('extractData() Error')

		# dataToLoad = True # hard set for testing 
		if dataToLoad:
			# Load Data
			try:
				self.loadData()
			except Exception as e:
				self.logger.exception('loadData() Error')
				self.StatusEmail('loadData() Error', e, traceback.format_exc())
				sys.exit('loadData() Error')

			# Process Data
			try:
				self.processData()
			except Exception as e:
				self.logger.exception('processData() Error')
				self.StatusEmail('processData() Error', e, traceback.format_exc())
				sys.exit('processData() Error')

			# Data Checks
			try:
				self.dataChecks()
			except Exception as e:
				self.logger.exception('dataChecks() Error')
				self.StatusEmail('dataChecks() Error', e, traceback.format_exc())
				sys.exit('dataChecks() Error')

if __name__ == '__main__':

	bh = ETL('Bahrain')
	bh.run()
