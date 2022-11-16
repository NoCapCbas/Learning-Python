import logging
from urllib import request 
import pandas as pd
import pyodbc
from datetime import datetime
import os
import traceback
from bs4 import BeautifulSoup as bs
import sys
from time import sleep
import shutil
import zipfile
import numpy as np
import glob
from pywintypes import com_error
from pathlib import Path
from pandasql import sqldf
import requests 
from bs4 import BeautifulSoup as bs
import zipfile
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
		self.driver = 0
		conn = pyodbc.connect(
			'Driver={SQL Server};'
			'Server=SEVENFARMS_DB3;'
			'Database=Control;'
			'UID=sa;'
			'PWD=Harpua88;'
			'Trusted_Connection=No;'
		)
		stopYM_Q = f'''
		SELECT StopYM
		FROM [Control].[dbo].[Data_Availability_Monthly]
		WHERE Declarant = '{self.country}'
		'''
		
		self.lastPeriodOfTDMdata = list(pd.read_sql_query(stopYM_Q, conn)['StopYM'])[0]
		# print(f'TDM Latest Published Period: {self.lastPeriodOfTDMdata}')
		if int(self.lastPeriodOfTDMdata[-2:]) == 12:
			self.yearToQuery = str(int(self.lastPeriodOfTDMdata[:-2]) + 1)
			self.nextMon = '01'
		else:
			self.yearToQuery = self.lastPeriodOfTDMdata[:-2]
			self.nextMon = str(int(self.lastPeriodOfTDMdata[-2:]) + 1)
			if len(self.nextMon) == 1:
				self.nextMon = '0' + self.nextMon
		# print(self.yearToQuery)
		self.nextPeriodOfTDMdata = self.yearToQuery + str(self.nextMon)
		#self.nextPeriodOfTDMdata = '202205' # hard set for testing
		# print(f'Period to search for: {self.nextPeriodOfTDMdata}')
		
		conn.close()
		# URLS
		self.data_source_url = 'https://statsmauritius.govmu.org/Pages/Statistics/By_Subject/External_Trade/Detailed_Trade_Data.aspx' 
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
		self.logger.info('Mauritius_ETL')

	def StatusEmail(self, phase, text, text2 = ''):
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
		password = 'drwvgldfexinkmyc'
		message = MIMEMultipart("alternative")
		message["Subject"] = f'{self.country} {phase}'
		message["from"] = "TDM Data Team"
		message["To"] = ", ".join(recipients)

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

		wb.Close(True)
		excel.Quit()

	def autoUnitCheck(self, conn, cursor):
		self.logger.info('\t\tWorking on Unit Checks...')
		if not os.path.exists(rf'Y:\{self.country}\Unit Checks\{self.nextPeriodOfTDMdata}'):
			os.mkdir(rf'Y:\{self.country}\Unit Checks\{self.nextPeriodOfTDMdata}')

		f_path = Path(rf'Y:\{self.country}\Unit Checks\{self.nextPeriodOfTDMdata}')
		
		table_matches = {
			'E':'TEMP_STEP7_EXP',
			'I':'TEMP_STEP7_IMP',
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
			FROM [SRC_Mauritius].[dbo].[{k}8]
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
		self.currentDate = datetime.now().strftime('%d%m%y')
		#self.currentDate = '270722' # hard set for testing
		# print(self.currentDate)
		mainLink = 'https://statsmauritius.govmu.org/'
		
		# grabs mauritius data source html to grab direct download link
		r = requests.get(self.data_source_url)
		# print(r.content)
		# html to bs obj
		soup = bs(r.text)
		# grabs first table on site
		download_table = soup.find_all('table')[0]
		# print(download_table)

		# grab first non header row
		second_tr = download_table.find_all('tr')[1]
		# print(second_tr)

		# grab exp and imp anchors
		first2_anchors_in_first_row = second_tr.find_all('a')[:2]
		# print(all_anchors_in_first_row)

		# both exports and imports download link must match the current date to continue to loadData
		for a in first2_anchors_in_first_row:
			endExcelLink = a['href']
			filename = endExcelLink.split('/')[-1]
			date_in_filename = filename.split('.')[0].split('_')[-1]
			# if currentDate matches date in filename download files
			if date_in_filename == self.currentDate:
				self.logger.info(f'\t\tSite Filename({date_in_filename}) matches Current Date({self.currentDate})')
				# download files
				downloadLink = mainLink + endExcelLink
				self.logger.info(f'\t\t\tDownloading {downloadLink}...')
				r = requests.get(downloadLink)
				# write files to excel
				open(f'{self.downloadPath}\\{filename}', 'wb').write(r.content)
				sleep(5)
				# read file and confirm
				tempDF = pd.read_excel(f'{self.downloadPath}\\{filename}', skiprows=[0])
				# confirm new data is available
				self.logger.info(f'\tConfirming Data is New...')
				if self.nextPeriodOfTDMdata in list(tempDF[tempDF.columns[0]].str.replace('-', '')):
					self.logger.info('\t\t\t*New Data Available*')
					# convert excel to txt file
					tempDF.to_csv(f'{self.datafilesPath}\\{filename.split(".")[0]}.csv', encoding='utf-8', index=False)
					# move excel file to archive
					shutil.move(f'{self.downloadPath}\\{filename}', f'{self.archivePath}\\{filename}')
					self.StatusEmail('New Data', f'New {self.country} Data', '')
					dataToLoad = True
				else:
					dataToLoad = False
					self.logger.info('\t\t\t*Data did not pass*')
			else:
				dataToLoad = False
				self.logger.info(f'\t\tSite Filename({date_in_filename}) does not match Current Date({self.currentDate})')


		if not dataToLoad:
			self.logger.info('\t\tNo New Data is Available.')
			self.logger.info('\tData Extraction Terminated.')
			return dataToLoad

		self.logger.info('\tData Extraction Complete.')
		return dataToLoad
	
	def loadData(self):
		self.logger.info('\tLoading Data...')

		# create connection to [SRC_Mauritius].[dbo].[TEMP]
		conn = pyodbc.connect(
			driver = '{ODBC Driver 17 for SQL Server}', 
			server = 'SEVENFARMS_DB1', 
			database = 'SRC_Mauritius', 
			uid = 'sa', 
			pwd = 'Harpua88',
			)
		cursor = conn.cursor()

		MASTER = pd.DataFrame()
		# read all txt files in datafilesPath
		for file in os.listdir(self.datafilesPath):
			# if files is a csv
			if file.split('.')[-1] == 'csv':
				# read file
				tempDF = pd.read_csv(f'{self.datafilesPath}\\{file}')
				tempDF.insert(0, 'FILENAME', f'{self.datafilesPath}\\{file}')
				# add to master df
				MASTER = pd.concat([MASTER, tempDF], ignore_index=True)
				# move read file to archive
				shutil.move(f'{self.datafilesPath}\\{file}', f'{self.archivePath}\\{file}')
		# turn master datatypes to string
		MASTER = MASTER.astype(str)
		MASTER = MASTER.rename(columns={
			MASTER.columns[1]: 'TRANSACTION_MONTH',
			MASTER.columns[2]: 'CONTINENT',
			MASTER.columns[3]: 'COUNTRY',
			MASTER.columns[4]: 'MODE_OF_TRANSPORT',
			MASTER.columns[5]: 'SITC',
			MASTER.columns[6]: 'HS_CODE',
			MASTER.columns[7]: 'DESCRIPTION_OF_GOODS',
			MASTER.columns[8]: 'UNIT',
			MASTER.columns[9]: 'QUANTITY',
			MASTER.columns[10]: 'FREE_ON_BOARD',
			MASTER.columns[11]: 'COST_INSURANCE_FREIGHT',
		})
		# print(MASTER)
		# print(MASTER.columns)
		# load MASTER to DB
		self.logger.info('\t\t\tDROPPING TABLE [SRC_Mauritius].[dbo].[TEMP]')
		cursor.execute(f"IF OBJECT_ID('[SRC_Mauritius].[dbo].[TEMP]') IS NOT NULL DROP TABLE [SRC_Mauritius].[dbo].[TEMP];")
		cursor.commit()

		# CREATE TABLE
		create_statements_cols = ''
		for col in MASTER.columns: 
			if MASTER.columns[-1] == col:
				create_statements_cols = create_statements_cols + f'[{col}] [varchar](max) NULL'
			else:
				create_statements_cols = create_statements_cols + f'[{col}] [varchar](max) NULL,'
		# print(create_statements_cols)

		self.logger.info(f'\t\t\tCREATING TABLE [SRC_Mauritius].[dbo].[TEMP]')
		cursor.execute(f"""
		CREATE TABLE [SRC_Mauritius].[dbo].[TEMP](
			{create_statements_cols}
		);
		""")
		cursor.commit()


		# Loop through list of split df and insert into table
		self.logger.info(f'\t\t\tINSERTING {len(MASTER)} ROWS INTO SRC_Mauritius.dbo.TEMP')
		insert_to_temp_table = f'INSERT INTO [SRC_Mauritius].[dbo].[TEMP] VALUES (?,?,?,?,?,?,?,?,?,?,?,?)'
		cursor.fast_executemany = True
		
		# INSERT INTO TEMP
		masterValues = MASTER.values.tolist()
		# print(masterValues)
		cursor.executemany(insert_to_temp_table, masterValues)
		cursor.commit()

		cursor.close()
		conn.close()
		self.logger.info('\tData Loaded.')
	
	def processData(self):
		self.logger.info('\tProcessing Data...')
		# Grab SQL processing script 
		sqlScriptPath = "Y:\\_PyScripts\\Damon\\Mauritius\\Mauritius_pyscript.sql"
		sql = self.getSql(sqlScriptPath)
		# Create Connection
		conn = pyodbc.connect(
			driver = '{ODBC Driver 17 for SQL Server}', 
			server = 'SEVENFARMS_DB1', 
			database = 'SRC_Mauritius', 
			uid = 'sa', 
			pwd = 'Harpua88')

		cursor = conn.cursor()
		# Execute the read out sql script string.
		cursor.execute(sql)
		cursor.commit()
		running = True
		while running == True: 
			sleep(3)
			try: 
				query = cursor.execute('''
				SELECT * 
				FROM [SRC_Mauritius].[DBO].[RUNNING-STATUS];
				''').fetchone()

			except: 
				continue
			STATUS = query[0]
			if STATUS == 0: 
				self.logger.info('Data Processed.')
				running = False
				cursor.execute('''
				IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[SRC_Mauritius].[dbo].[RUNNING-STATUS]') AND type in (N'U'))
				DROP TABLE [SRC_Mauritius].[DBO].[RUNNING-STATUS];
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
			server = 'SEVENFARMS_DB1', 
			database = 'SRC_Mauritius', 
			uid = 'sa', 
			pwd = 'Harpua88')
		cursor = conn.cursor()

		# SQL script table names
		EXP_TABLE_NAME = '[SRC_Mauritius].[dbo].[TEMP_STEP7_EXP]'
		IMP_TABLE_NAME = '[SRC_Mauritius].[dbo].[TEMP_STEP7_IMP]'
		# Generic checks
		self.logger.info('\tGeneric Checks...')
		dfE8 = pd.read_sql_query(f'''
		SELECT *
		FROM {EXP_TABLE_NAME}
		''', conn)
		dfI8 = pd.read_sql_query(f'''
		SELECT *
		FROM {IMP_TABLE_NAME}
		''', conn)

		failedChecks = {'CHECK PERFORMED ON FINAL TABLE': ['STATUS'],
						'E NULL VALUES DO NOT EXIST':[False], 
						'I NULL VALUES DO NOT EXIST':[False], 
						'E COMMODITY_LEN IS 8 DIGITS':[False],
						'I COMMODITY_LEN IS 8 DIGITS':[False],
						'ALL E CTY_PTN EXIST IN CTY_MASTER DB1': [False],
						'ALL I CTY_PTN EXIST IN CTY_MASTER DB1': [False],
						'ALL E UNIT1 EXIST IN UOM_MASTER DB1': [False],
						'ALL I UNIT1 EXIST IN UOM_MASTER DB1': [False],
						'ALL E UNIT2 EXIST IN UOM_MASTER DB1': [False],
						'ALL I UNIT2 EXIST IN UOM_MASTER DB1': [False],
						'E ROW COUNT': [len(dfE8)],
						'I ROW COUNT': [len(dfI8)],
		}

		if dfE8.isnull().values.any() != True:
			failedChecks['E NULL VALUES DO NOT EXIST'][0] = True

		if dfI8.isnull().values.any() != True:
			failedChecks['I NULL VALUES DO NOT EXIST'][0] = True

		if dfE8['COMMODITY'].astype(str).map(len).nunique() == 1 and dfE8['COMMODITY'].astype(str).map(len).unique()[0] == 8:
			failedChecks['E COMMODITY_LEN IS 8 DIGITS'][0] = True

		if dfI8['COMMODITY'].astype(str).map(len).nunique() == 1 and dfI8['COMMODITY'].astype(str).map(len).unique()[0] == 8:
			failedChecks['I COMMODITY_LEN IS 8 DIGITS'][0] = True
		# Check CTY_PTN
		self.logger.info('\t\tChecking CTY_PTN...')
		tempE8 = pd.read_sql_query(f'''
		SELECT DISTINCT CTY_PTN, CTY_ABBR
		FROM {EXP_TABLE_NAME} A
		LEFT JOIN [SP_MASTER].[dbo].[CTY_MASTER] B
		ON A.CTY_PTN = B.CTY_ABBR
		''', conn)
		if tempE8.isnull().values.any() != True: 
			failedChecks['ALL E CTY_PTN EXIST IN CTY_MASTER DB1'][0] = True

		tempI8 = pd.read_sql_query(f'''
		SELECT DISTINCT CTY_PTN, CTY_ABBR
		FROM {IMP_TABLE_NAME} A
		LEFT JOIN [SP_MASTER].[dbo].[CTY_MASTER] B
		ON A.CTY_PTN = B.CTY_ABBR
		''', conn)
		if tempI8.isnull().values.any() != True: 
			failedChecks['ALL I CTY_PTN EXIST IN CTY_MASTER DB1'][0] = True
		
		# Check UNIT1
		self.logger.info('\t\tChecking UNIT1...')
		tempE8 = pd.read_sql_query(f'''
		SELECT DISTINCT UNIT1, ANH_UOM
		FROM {EXP_TABLE_NAME} A
		LEFT JOIN [SP_MASTER].[dbo].[UOM_MASTER] B
		ON A.UNIT1 = B.ANH_UOM
		''', conn)
		if tempE8.isnull().values.any() != True: 
			failedChecks['ALL E UNIT1 EXIST IN UOM_MASTER DB1'][0] = True

		tempI8 = pd.read_sql_query(f'''
		SELECT DISTINCT UNIT1, ANH_UOM
		FROM {IMP_TABLE_NAME} A
		LEFT JOIN [SP_MASTER].[dbo].[UOM_MASTER] B
		ON A.UNIT1 = B.ANH_UOM
		''', conn)
		if tempI8.isnull().values.any() != True: 
			failedChecks['ALL I UNIT1 EXIST IN UOM_MASTER DB1'][0] = True
		
		# Check UNIT2
		self.logger.info('\t\tChecking UNIT2...')
		tempE8 = pd.read_sql_query(f'''
		SELECT DISTINCT UNIT2, ANH_UOM
		FROM {EXP_TABLE_NAME} A
		LEFT JOIN [SP_MASTER].[dbo].[UOM_MASTER] B
		ON A.UNIT2 = B.ANH_UOM
		''', conn)
		if tempE8.isnull().values.any() != True: 
			failedChecks['ALL E UNIT2 EXIST IN UOM_MASTER DB1'][0] = True

		tempI8 = pd.read_sql_query(f'''
		SELECT DISTINCT UNIT2, ANH_UOM
		FROM {IMP_TABLE_NAME} A
		LEFT JOIN [SP_MASTER].[dbo].[UOM_MASTER] B
		ON A.UNIT2 = B.ANH_UOM
		''', conn)
		if tempI8.isnull().values.any() != True: 
			failedChecks['ALL I UNIT2 EXIST IN UOM_MASTER DB1'][0] = True
		
		# print(failedChecks)
		genericChecks = pd.DataFrame.from_dict(failedChecks)
		genericChecks = genericChecks.set_index('CHECK PERFORMED ON FINAL TABLE').transpose()

		# TOTAL CHECKS
		self.logger.info('\t\tQuerying Totals...')
		EXP_TOTALS = pd.read_sql_query('''
		SELECT CAST(SUM([VALUE_RUPEE]) AS VARCHAR) AS [PROCESSED TOTAL]
		FROM [SRC_Mauritius].[dbo].[TEMP_STEP7_EXP]
		''',conn)
		IMP_TOTALS = pd.read_sql_query('''
		SELECT CAST(SUM([VALUE_FOB_RUPEE]) AS VARCHAR) AS [PROCESSED FOB TOTAL]
			,CAST(SUM([VALUE_CIF_RUPEE]) AS VARCHAR) AS [PROCESSED CIF TOTAL]
		FROM [SRC_Mauritius].[dbo].[TEMP_STEP7_IMP]
		''',conn)

		tempDF = pd.read_sql_query('''
		SELECT *
		FROM [SRC_Mauritius].[dbo].[TEMP]
		''',conn)
		mask = tempDF[tempDF.isin(['TOTAL', 'TOTAL ', ' TOTAL', ' TOTAL ']).any(axis=1)]
		TOTALS_LIST = mask.values.tolist()
		for row in TOTALS_LIST: 
			try:
				while True:
					row.remove('nan')
			except:
				pass

			if 'Exp' in row[0]:
				EXP_TOTALS.insert(0, 'RAW TOTAL', row[-1])
			if 'Imp' in row[0]:
				IMP_TOTALS.insert(0, 'RAW CIF TOTAL', row[-1])
				IMP_TOTALS.insert(0, 'RAW FOB TOTAL', row[-2])
		# print(EXP_TOTALS)
		# print(IMP_TOTALS)

		# Check if newest period exists within final tables before inserting, want a table of length zero 
		check_E8 = pd.read_sql_query(f"""
		SELECT DISTINCT a.PERIOD, b.PERIOD
		FROM {EXP_TABLE_NAME} a
		LEFT JOIN [SRC_Mauritius].[dbo].[E8] b
		ON a.PERIOD = b.PERIOD
		WHERE b.PERIOD IS NOT NULL
		AND a.PERIOD = '{self.nextPeriodOfTDMdata}';
		""", conn)
		# print(len(check_E8))
		check_I8 = pd.read_sql_query(f"""
		SELECT DISTINCT a.PERIOD, b.PERIOD
		FROM {IMP_TABLE_NAME} a
		LEFT JOIN [SRC_Mauritius].[dbo].[I8] b
		ON a.PERIOD = b.PERIOD
		WHERE b.PERIOD IS NOT NULL
		AND a.PERIOD = '{self.nextPeriodOfTDMdata}';
		""", conn)
		# print(len(check_I8))

		insrtCondition = False
		if list(genericChecks.all())[0] and len(check_E8) == 0 and len(check_I8) == 0:
			# INSERT INTO ARCHIVE 
			cursor.execute("""
			-- DELETE OLD DATA IN ARCHIVE IF PERIOD MATCHES
			DELETE FROM [SRC_Mauritius].[dbo].[SRC_EXPORTS_ARCHIVE]
			WHERE PERIOD IN (
			SELECT DISTINCT [TRANSACTION_MONTH]
			FROM [SRC_Mauritius].[dbo].[TEMP_STEP0]
			WHERE [FILENAME] LIKE '%Exports%');

			DELETE FROM [SRC_Mauritius].[dbo].[SRC_IMPORTS_ARCHIVE]
			WHERE PERIOD IN (
			SELECT DISTINCT [TRANSACTION_MONTH]
			FROM [SRC_Mauritius].[dbo].[TEMP_STEP0]
			WHERE [FILENAME] LIKE '%Imports%');

		-- EXPORTS 
			INSERT INTO [SRC_Mauritius].[dbo].[SRC_EXPORTS_ARCHIVE]
			SELECT [FILENAME]
			,[TRANSACTION_MONTH] AS [PERIOD]
			,SUBSTRING([TRANSACTION_MONTH], 1,4) AS [Year]
			,SUBSTRING([TRANSACTION_MONTH], 5,6) AS [Mon]
			,[HS_CODE] AS [HS Code]
			,REPLACE([SITC], '.0', '') AS [SITC]
			,[DESCRIPTION_OF_GOODS] AS [Description of goods]
			,[COUNTRY] AS [Country]
			,[UNIT] AS [Unit]
			,[MODE_OF_TRANSPORT] AS [Mode of transport]
			,[QUANTITY] AS [Quantity]
			,[FREE_ON_BOARD] AS [Free on board (in Mauritian Rupee)]
			FROM [SRC_Mauritius].[dbo].[TEMP_STEP0]
			WHERE [FILENAME] LIKE '%Exports%';
		-- IMPORTS
			INSERT INTO [SRC_Mauritius].[dbo].[SRC_IMPORTS_ARCHIVE]
			SELECT [FILENAME]
			,[TRANSACTION_MONTH] AS [PERIOD]
			,SUBSTRING([TRANSACTION_MONTH], 1,4) AS [Year]
			,SUBSTRING([TRANSACTION_MONTH], 5,6) AS [Mon]
			,[HS_CODE] AS [HS Code]
			,REPLACE([SITC], '.0', '') AS [SITC]
			,[DESCRIPTION_OF_GOODS] AS [Description of goods]
			,[COUNTRY] AS [Country]
			,[UNIT] AS [Unit]
			,[MODE_OF_TRANSPORT] AS [Mode of transport]
			,[QUANTITY] AS [Quantity]
			,[FREE_ON_BOARD] AS [Free on board (in Mauritian Rupee)]
			,[COST_INSURANCE_FREIGHT] AS [Cost insurance freight (in Mauritian Rupee)]
			FROM [SRC_Mauritius].[dbo].[TEMP_STEP0]
			WHERE [FILENAME] LIKE '%Imports%';
			""")
			cursor.commit()

			# insert into final tables
			cursor.execute('''
			-- EXPORTS 
			DELETE FROM [SRC_Mauritius].[dbo].[E8] 
			WHERE PERIOD IN (
				SELECT DISTINCT PERIOD 
				FROM [SRC_Mauritius].[dbo].[TEMP_STEP7_EXP]
			);
			INSERT INTO [SRC_Mauritius].[dbo].[E8]
			SELECT cast([CTY_RPT] AS VARCHAR(2)) AS [CTY_RPT]
				,cast([CTY_PTN] AS VARCHAR(7)) AS [CTY_PTN]
				,cast(replace ([COMMODITY],' ','') AS VARCHAR(12)) AS [COMMODITY]
				,cast([PERIOD] AS VARCHAR(6)) AS [PERIOD]
				,cast([YR] AS INT) AS [YR]
				,cast([MON] AS TINYINT) AS [MON]
				,cast([VALUE_RUPEE] AS DECIMAL(38,8)) AS [VALUE_RUPEE]
				,cast([VALUE] AS DECIMAL(38,8)) AS [VALUE]
				,cast([UNIT1] AS VARCHAR(3)) AS [UNIT1]
				,cast([QTY1] AS DECIMAL(18,2)) AS [QTY1]
				,cast([UNIT2] AS VARCHAR(3)) AS [UNIT2]
				,cast([QTY2] AS DECIMAL(18,2)) AS [QTY2]
			FROM [SRC_Mauritius].[dbo].[TEMP_STEP7_EXP];

			-- IMPORTS
			DELETE FROM [SRC_Mauritius].[dbo].[I8] 
			WHERE PERIOD IN (
				SELECT DISTINCT PERIOD 
				FROM [SRC_Mauritius].[dbo].[TEMP_STEP7_IMP]
			);
			INSERT INTO [SRC_Mauritius].[dbo].[I8] 
			SELECT  cast([CTY_RPT] AS VARCHAR(2)) AS [CTY_RPT]
				,cast([CTY_PTN] AS VARCHAR(7)) AS [CTY_PTN]
				,cast(replace ([COMMODITY],' ','') AS VARCHAR(12)) AS [COMMODITY]
				,cast([PERIOD] AS VARCHAR(6)) AS [PERIOD]
				,cast([YR] AS INT) AS [YR]
				,cast([MON] AS TINYINT) AS [MON]
				,cast([VALUE_FOB_RUPEE] AS DECIMAL(38,8)) AS [VALUE_RUPEE]
				,cast([VALUE_FOB] AS DECIMAL(38,8)) AS [VALUE]
				,cast([UNIT1] AS VARCHAR(3)) AS [UNIT1]
				,cast([QTY1] AS DECIMAL(18,2)) AS [QTY1]
				,cast([UNIT2] AS VARCHAR(3)) AS [UNIT2]
				,cast([QTY2] AS DECIMAL(18,2)) AS [QTY2]
			FROM [SRC_Mauritius].[dbo].[TEMP_STEP7_IMP];

			-- IMPORTS FOB CIF
			DELETE FROM [SRC_Mauritius].[dbo].[I8_FOB_CIF] 
			WHERE PERIOD IN (
				SELECT DISTINCT PERIOD 
				FROM [SRC_Mauritius].[dbo].[TEMP_STEP7_IMP]
			);
			INSERT INTO [SRC_Mauritius].[dbo].[I8_FOB_CIF] 
			SELECT  cast([CTY_RPT] AS VARCHAR(2)) AS [CTY_RPT]
				,cast([CTY_PTN] AS VARCHAR(7)) AS [CTY_PTN]
				,cast(replace ([COMMODITY],' ','') AS VARCHAR(12)) AS [COMMODITY]
				,cast([PERIOD] AS VARCHAR(6)) AS [PERIOD]
				,cast([YR] AS INT) AS [YR]
				,cast([MON] AS TINYINT) AS [MON]
				,cast([VALUE_FOB_RUPEE] AS DECIMAL(38,8)) AS [VALUE_FOB_RUPEE]
				,cast([VALUE_FOB] AS DECIMAL(38,8)) AS [VALUE_FOB]
				,cast([VALUE_CIF_RUPEE] AS DECIMAL(38,8)) AS [VALUE_CIF_RUPEE]
				,cast([VALUE_CIF] AS DECIMAL(38,8)) AS [VALUE_CIF]
				,cast([UNIT1] AS VARCHAR(3)) AS [UNIT1]
				,cast([QTY1] AS DECIMAL(18,2)) AS [QTY1]
				,cast([UNIT2] AS VARCHAR(3)) AS [UNIT2]
				,cast([QTY2] AS DECIMAL(18,2)) AS [QTY2]
			FROM [SRC_Mauritius].[dbo].[TEMP_STEP7_IMP];

			-- IMPORTS I8 CIF
			DELETE FROM [SRC_Mauritius].[dbo].[I8_CIF] 
			WHERE PERIOD IN (
				SELECT DISTINCT PERIOD 
				FROM [SRC_Mauritius].[dbo].[TEMP_STEP7_IMP]
			);
			INSERT INTO [SRC_Mauritius].[dbo].[I8_CIF] 
			SELECT  'MUC' AS [CTY_RPT]
				,cast([CTY_PTN] AS VARCHAR(7)) AS [CTY_PTN]
				,cast(replace ([COMMODITY],' ','') AS VARCHAR(12)) AS [COMMODITY]
				,cast([PERIOD] AS VARCHAR(6)) AS [PERIOD]
				,cast([YR] AS INT) AS [YR]
				,cast([MON] AS TINYINT) AS [MON]
				,cast([VALUE_CIF_RUPEE] AS DECIMAL(38,8)) AS [VALUE_RUPEE]
				,cast([VALUE_CIF] AS DECIMAL(38,8)) AS [VALUE]
				,cast([UNIT1] AS VARCHAR(3)) AS [UNIT1]
				,cast([QTY1] AS DECIMAL(18,2)) AS [QTY1]
				,cast([UNIT2] AS VARCHAR(3)) AS [UNIT2]
				,cast([QTY2] AS DECIMAL(18,2)) AS [QTY2]
			FROM [SRC_Mauritius].[dbo].[TEMP_STEP7_IMP];

			''')
			cursor.commit()
			# drop tables
			cursor.execute('''
			--DROP TABLE [SRC_Mauritius].[dbo].[TEMP];
			DROP TABLE [SRC_Mauritius].[dbo].[TEMP_STEP0];
			DROP TABLE [SRC_Mauritius].[dbo].[TEMP_STEP1];
			DROP TABLE [SRC_Mauritius].[dbo].[TEMP_STEP2];
			DROP TABLE [SRC_Mauritius].[dbo].[TEMP_STEP3];
			DROP TABLE [SRC_Mauritius].[dbo].[TEMP_STEP4];
			DROP TABLE [SRC_Mauritius].[dbo].[TEMP_STEP5];
			DROP TABLE [SRC_Mauritius].[dbo].[TEMP_STEP6];
			--DROP TABLE [SRC_Mauritius].[dbo].[TEMP_STEP7_EXP];
			--DROP TABLE [SRC_Mauritius].[dbo].[TEMP_STEP7_IMP];
			''')
			cursor.commit()
			
			# Unit Checks
			try:
				conn, cursor = self.autoUnitCheck(conn, cursor)
			except Exception as e:
				self.StatusEmail('autoUnitCheck() Error', e, traceback.format_exc())
			insrtCondition = True
			self.logger.info('Data Verified.')
		else:
			self.logger.info(f'INSERT STATEMENT EXECUTED: {insrtCondition}')
		
		# Send Check Email
		port = 465
		recipients = ["DDiaz@tradedatamonitor.com"
					# ,"y.zeng@tradedatamonitor.com"
					# ,"a.chan@tradedatamonitor.com"
					,"j.smith@tradedatamonitor.com"
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
		self.logger.info(f'Launching {self.country} ETL')

		# Data Extraction
		try:
			dataToLoad = self.extractData()
		except Exception as e:
			self.logger.exception('extractData() Error')
			self.StatusEmail('extractData() Error', e, traceback.format_exc())
			sys.exit()

		#dataToLoad = True # hard set for testing
		if dataToLoad:
			# Load Data
			try:
				self.loadData()
			except Exception as e:
				self.logger.exception('loadData() Error')
				self.StatusEmail('loadData() Error', e, traceback.format_exc())
				sys.exit()
			
			# Process Data
			try:
				self.processData()
			except Exception as e:
				self.logger.exception('processData() Error')
				self.StatusEmail('processData() Error', e, traceback.format_exc())
				sys.exit()

			# Data Checks
			try:
				self.dataChecks()
			except Exception as e:
				self.logger.exception('dataChecks() Error')
				self.StatusEmail('dataChecks() Error', e, traceback.format_exc())
				sys.exit()
		

		self.logger.info(f'{self.country} ETL Complete.')


if __name__ == '__main__':

	mu = ETL('Mauritius')
	mu.run()
