import logging
from sqlite3 import connect
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
			'Server=SF-TEST\SFTESTDB;'
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
		self.excelBugPath = 'C:\\Users\\DDiaz.ANH\\AppData\\Local\\Temp\\gen_py\\3.10\\00020813-0000-0000-C000-000000000046x0x1x9'
		if os.path.exists(self.excelBugPath):
			shutil.rmtree(self.excelBugPath)
		
		self.lastPeriodOfTDMdata = list(pd.read_sql_query(stopYM_Q, conn)['StopYM'])[0]
		# print(self.lastPeriodOfTDMdata)
		self.StopY = self.lastPeriodOfTDMdata[:-2]
		self.StopMon = self.lastPeriodOfTDMdata[-2:]
		if int(self.StopMon) == 12:
			self.nextY = str(int(self.lastPeriodOfTDMdata[:-2]) + 1)
			self.nextMon = '01'
		else:
			self.nextY = str(int(self.lastPeriodOfTDMdata[:-2]))
			self.nextMon = str(int(self.StopMon) + 1)
			if len(self.nextMon) == 1:
				self.nextMon = '0' + self.nextMon

		self.nextPeriodOfTDMdata = self.nextY + self.nextMon
		# self.nextPeriodOfTDMdata = '202205' # testing
		conn.close()
		# URLS
		self.data_source_url = 'https://lekiosque.finances.gouv.fr/site_fr/telechargement/telechargement_SGBD.asp' 
		# PATHS
		self.logPath = f"Y:\\_PyScripts\\Damon\\{self.country.replace(' ', '_')}\\Log"
		self.downloadPath = f"Y:\\_PyScripts\\Damon\\{self.country.replace(' ', '_')}\\Downloads"
		self.datafilesPath = f'Y:\\{self.country.replace(" ", "_")}\\Data Files'
		self.archivePath = f'Y:\\{self.country.replace(" ", "_")}\\Archive\\{self.nextPeriodOfTDMdata}'
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
		self.logger.info('France_Customs_ETL')
		self.logger.info(f'Last Period Published: {self.lastPeriodOfTDMdata}')

	def StatusEmail(self, phase, text, text2 = '', ALL=False):
		# Grab RUNTIME.log
		with open(self.logPath + "\\RUNTIME.log", mode='r') as fileObj:
			RUNTIMElog = fileObj.read()

		port = 465
		smtp_server = "smtp.gmail.com"
		sender_email = "tdmUsageAlert@gmail.com"
		if ALL:
			recipients = ["DDiaz@tradedatamonitor.com",
						"tdmdata@tradedatamonitor.com"
						,"m.alomar@tradedatamonitor.com"
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
	
	def connect(self, URL):
		self.logger.info(f'\t\tConnecting {URL}')

		try:
			chrome_options = webdriver.ChromeOptions()
			prefs = {'download.default_directory':self.downloadPath}
			chrome_options.add_experimental_option('prefs', prefs)
			self.driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=chrome_options)
			self.driver.get(URL)
		except Exception as e:
			self.logger.info(traceback.format_exc())
			if self.driver != 0 and self.driver.session_id is not None:
				self.driver.quit()
				self.driver = 0
			self.connect(URL)

	def getSql(self, sqlPath):
		# Open the external sql file.
		file = open(sqlPath, 'r')
		# Read out the sql script text in the file.
		sql = file.read()
		# Close the sql file object.
		file.close()
		return sql

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
		if not os.path.exists(rf'Y:\{self.country.replace(" ", "_")}\Unit Checks\{self.nextPeriodOfTDMdata}'):
			os.mkdir(rf'Y:\{self.country.replace(" ", "_")}\Unit Checks\{self.nextPeriodOfTDMdata}')

		f_path = Path(rf'Y:\{self.country.replace(" ", "_")}\Unit Checks\{self.nextPeriodOfTDMdata}')
		
		table_matches = {
			'E':'TEMP_STEP6_EXP',
			'I':'TEMP_STEP6_IMP',
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
			FROM [SRC_France_Customs].[dbo].[{k}8]
			GROUP BY PERIOD,unit1,SUBSTRING(COMMODITY,1,2) ,unit2,cty_rpt
			ORDER BY period
			''', conn)
			
			pivotData.to_excel(f_path/file_name, sheet_name=f'{k}8', index=False)
			self.generateUnitChecks(f_path, file_name, f'{k}8')
			if os.path.exists(self.excelBugPath):
				shutil.rmtree(self.excelBugPath)
		
		self.logger.info('\t\t\tUnit Checks Generated.')
		
		return conn, cursor

	def extractData(self):
		self.logger.info('\tExtracting Data...')
		self.currentDate = datetime.now().strftime('%Y-%m/%d')
		# self.currentDate = '2022-07/08' 
		urls_to_scrape = []
		# grabs france data source html to grab direct download link
		r = requests.get(self.data_source_url)
		# direct download file
		# url = f'https://www.douane.gouv.fr/sites/default/files/{self.currentDate}/{self.lastPeriodOfTDMdata}-stat-national-ce-export.zip'
		
		soup = bs(r.text)
		html = soup.find_all('div', {'id':'gauche'})[0]
		# print(html)
		file_sections = html.find_all('div', {'class':'bande'})
		for div in file_sections:
			url = div.find_all('a')[0]['href']
			# print(url)
			
			url_split = url.split('.')
			# add url to urls_to_scrape if it meets requirements
			if len(url_split) > 2 and ('National' in url or 'national' in url) and self.currentDate in url and self.nextPeriodOfTDMdata in url:
				# print(url)
				urls_to_scrape.append(url)
			else:
				continue

		if urls_to_scrape:
			self.logger.info('\t\t*New Data Available*')
			dataToLoad = True
			self.StatusEmail('New Data', f'New {self.country} Data Downloading', '')
			# download file from each url scraped
			for url in urls_to_scrape:
				self.logger.info(f'\t\t\tDownloading {url}')
				r = requests.get(url)
				# print(f'{self.downloadPath}\\{url.split("/")[-1]}')
				open(f'{self.downloadPath}\\{url.split("/")[-1]}','wb').write(r.content)
		else:
			self.logger.info('\t\tNo New Data Available')
			dataToLoad = False

		self.logger.info('\tData Extraction Complete.')
		return dataToLoad
	
	def loadData(self):
		self.logger.info('\tLoading Data...')

		# create connection to [SRC_France_Customs].[dbo].[TEMP]
		conn = pyodbc.connect(
			driver = '{ODBC Driver 17 for SQL Server}', 
			server = 'SF-PROCESS', 
			database = 'SRC_France_Customs', 
			uid = 'sa', 
			pwd = 'Harpua88',
			)
		cursor = conn.cursor()

		dataFiles = []
		MASTER = pd.DataFrame()
		for file in os.listdir(self.downloadPath):
			
			zipfolderPath = f'{self.downloadPath}\\{file}'
			# print(zipfile.is_zipfile(zipfolderPath))
			
			with zipfile.ZipFile(zipfolderPath, 'r') as zipRef:
				zipRef.extractall(self.datafilesPath)
			
			innerFiles = os.listdir(self.datafilesPath + '\\' + file[7:-4])
			sofar = 0
			for innerfile in innerFiles:
				if os.path.isdir(self.datafilesPath + '\\' + file[7:-4]):
					# print(innerfile)
					size = os.path.getsize(self.datafilesPath + '\\' + file[7:-4] + '\\' + innerfile)
					if size > sofar:
							sofar = size
							max_file = innerfile
			# print(max_file)
			# read txt file
			temp = pd.read_csv(self.datafilesPath + '\\' + file[7:-4] + '\\' + max_file, sep=';', dtype=str, header=None)
			# add filename col to temp df
			temp.insert(0, 'FILENAME', self.datafilesPath + '\\' + file[7:-4] + '\\' + max_file)
			MASTER = pd.concat([MASTER, temp], ignore_index=True)
			# print(MASTER)
			# moves zip folder to archive once loaded
			shutil.move(f'{self.downloadPath}\\{file}', f'{self.archivePath}\\{file}')
		# moves all files extracted to datafilesPath to archive
		for file in os.listdir(self.datafilesPath):
			shutil.move(f'{self.datafilesPath}\\{file}',f'{self.archivePath}\\{file}')

		# print(MASTER.dtypes)
		MASTER = MASTER.astype(str)
		# print(MASTER.dtypes)
		MASTER = MASTER.rename(columns={
			MASTER.columns[1]: 'FLUX',
			MASTER.columns[2]: 'MDEP',
			MASTER.columns[3]: 'ADEP',
			MASTER.columns[4]: 'NC6',
			MASTER.columns[5]: 'SECTION',
			MASTER.columns[6]: 'NC8',
			MASTER.columns[7]: 'PYOD',
			MASTER.columns[8]: 'VART',
			MASTER.columns[9]: 'QUAN',
			MASTER.columns[10]: 'USUP',
		})
		# load MASTER to DB
		self.logger.info(f'\t\t\tDROPPING TABLE SRC_France_Customs.dbo.TEMP')
		cursor.execute(f"IF OBJECT_ID('SRC_France_Customs.dbo.TEMP') IS NOT NULL DROP TABLE [SRC_France_Customs].[dbo].[TEMP];")
		cursor.commit()
		
		# CREATE TABLE 
		create_statements_cols = ''
		for col in MASTER.columns: 
			if MASTER.columns[-1] == col:
				create_statements_cols = create_statements_cols + f'[{col}] [varchar](max) NULL'
			else:
				create_statements_cols = create_statements_cols + f'[{col}] [varchar](max) NULL,'
		

		self.logger.info(f'\t\t\tCREATING TABLE SRC_France_Customs.dbo.TEMP')
		cursor.execute(f"""
		CREATE TABLE SRC_France_Customs.dbo.TEMP(
			{create_statements_cols}
		)
		""")
		cursor.commit()

		# Loop through list of split df and insert into table
		self.logger.info(f'\t\t\tINSERTING {len(MASTER)} ROWS INTO SRC_France_Customs.dbo.TEMP')
		self.logger.info('')
		insert_to_temp_table = f'INSERT INTO SRC_France_Customs.dbo.TEMP VALUES (?,?,?,?,?,?,?,?,?,?,?)'
		cursor.fast_executemany = True

		# INSERT INTO TEMP
		masterValues = MASTER.values.tolist()
		rowStepper = 1000000
		# inserting 1000000 rows of data at a time to prevent error
		for rowI in range(0,len(masterValues), rowStepper):
			cursor.executemany(insert_to_temp_table, masterValues[rowI:rowI+rowStepper])
			cursor.commit()

		cursor.close()
		conn.close()
		self.logger.info('\tData Loaded.')
	
	def processData(self):
		self.logger.info('\tProcessing Data...')
		# Grab SQL processing script 
		sqlScriptPath = "Y:\\_PyScripts\\Damon\\France_Customs\\FRANCE_CUSTOMS_pyscript.sql"
		sql = self.getSql(sqlScriptPath)
		# Create Connection
		conn = pyodbc.connect(
			driver = '{ODBC Driver 17 for SQL Server}', 
			server = 'SF-PROCESS', 
			database = 'SRC_France_Customs', 
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
				FROM [SRC_France_Customs].[DBO].[RUNNING-STATUS];
				''').fetchone()

			except: 
				continue
			STATUS = query[0]
			if STATUS == 0: 
				self.logger.info('\t\t\tSQL Script Executed.')
				running = False
				cursor.execute('''
				IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[SRC_France_Customs].[dbo].[RUNNING-STATUS]') AND type in (N'U'))
				DROP TABLE [SRC_France_Customs].[DBO].[RUNNING-STATUS];
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
			database = 'SRC_France_Customs', 
			uid = 'sa', 
			pwd = 'Harpua88')
		cursor = conn.cursor()

		# SQL script table names
		EXP_TABLE_NAME = '[SRC_France_Customs].[dbo].[TEMP_STEP6_EXP]'
		IMP_TABLE_NAME = '[SRC_France_Customs].[dbo].[TEMP_STEP6_IMP]'

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
		# print(genericChecks)
		
		# TOTAL CHECKS
		self.logger.info('\t\tQuerying Totals...')
		EXP_TOTALS = pd.read_sql_query(f'''
		SELECT [CTY_RPT], [YR], CAST(SUM([VALUE]) AS VARCHAR) AS [VALUE_USD], CAST(SUM([VALUE_EURO]) AS VARCHAR) AS [VALUE_EURO]
		FROM {EXP_TABLE_NAME}
		GROUP BY [CTY_RPT], [YR]
		ORDER BY 2 ASC
  		''', conn)

		IMP_TOTALS = pd.read_sql_query(f'''
		SELECT [CTY_RPT], [YR], CAST(SUM([VALUE]) AS VARCHAR) AS [VALUE_USD], CAST(SUM([VALUE_EURO]) AS VARCHAR) AS [VALUE_EURO]
		FROM {IMP_TABLE_NAME}
		GROUP BY [CTY_RPT], [YR]
		ORDER BY 2 ASC
  		''', conn)

		# Check if newest period exists within final tables before inserting, want a table of length zero 
		check_E8 = pd.read_sql_query(f'''
		SELECT DISTINCT a.PERIOD, b.PERIOD
		FROM {EXP_TABLE_NAME} a
		LEFT JOIN [SRC_France_Customs].[dbo].[E8] b
		ON a.PERIOD = b.PERIOD
		WHERE b.PERIOD IS NOT NULL
		AND a.PERIOD = '{self.nextPeriodOfTDMdata}';
		''', conn)
		# print(len(check_E8))
		check_I8 = pd.read_sql_query(f'''
		SELECT DISTINCT a.PERIOD, b.PERIOD
		FROM {IMP_TABLE_NAME} a
		LEFT JOIN [SRC_France_Customs].[dbo].[I8] b
		ON a.PERIOD = b.PERIOD
		WHERE b.PERIOD IS NOT NULL
		AND a.PERIOD = '{self.nextPeriodOfTDMdata}';
		''', conn)
		# print(len(check_I8))

		insrtCondition = False
		if list(genericChecks.all())[0] and len(check_E8) == 0 and len(check_I8) == 0:
			# INSERT INTO ARCHIVE 
			exportTxtPath = 'Y:\\France_Customs\\Data Files\\stat-national-ce-export\\NATIONAL_NC8PAYSE.txt'
			importTxtPath = 'Y:\\France_Customs\\Data Files\\stat-national-ce-import\\NATIONAL_NC8PAYSI.txt'
			cursor.execute(f"""
			-- INSERT INTO MONTHLY ARCHIVE
				-- DELETE OLD DATA ARCHIVE IF YEAR, PERIOD MATCHES AND WHERE FILENAMES COME FROM 13 MONTH FILES
					DELETE FROM [SRC_France_Customs].[dbo].[ARCHIVE_EXP_IMP_NAT]
					WHERE CONCAT([ADEP],[MDEP]) IN (
					SELECT DISTINCT CONCAT([ADEP],[MDEP]) FROM [SRC_France_Customs].[dbo].[TEMP] 
					WHERE FILENAME = '{exportTxtPath}' 
					OR FILENAME = '{importTxtPath}' );

				-- INSERT DATA INTO ARCHIVE
					INSERT INTO [SRC_France_Customs].[dbo].[ARCHIVE_EXP_IMP_NAT]
					SELECT [FILENAME]
						,[FLUX]
						,[MDEP]
						,[ADEP]
						,[NC8]
						,[PYOD]
						,[VART]
						,[QUAN]
						,[USUP]
					FROM [SRC_France_Customs].[dbo].[TEMP]
					WHERE FILENAME = '{exportTxtPath}' 
					OR FILENAME = '{importTxtPath}';

				-- DELETE OLD YEARLY DATA ARCHIVE IF YEAR, PERIOD MATCHES AND WHERE FILENAMES DON'T COME FROM 13 MONTH FILES
					DELETE FROM [SRC_France_Customs].[dbo].[ARCHIVE_EXP_IMP_NAT_YEARLY]
					WHERE CONCAT([ADEP],[MDEP]) IN (
					SELECT DISTINCT CONCAT([ADEP],[MDEP]) FROM [SRC_France_Customs].[dbo].[TEMP] 
					WHERE [FILENAME] != '{exportTxtPath}' 
					OR [FILENAME] != '{importTxtPath}');

				-- INSERT DATA INTO ARCHIVE
					INSERT INTO [SRC_France_Customs].[dbo].[ARCHIVE_EXP_IMP_NAT_YEARLY]
					SELECT *
					FROM [SRC_France_Customs].[dbo].[TEMP]
					WHERE [FILENAME] != '{exportTxtPath}' 
					OR [FILENAME] != '{importTxtPath}';
					""")
			cursor.commit()
			# insert into final tables
			cursor.execute('''
			-- EXPORTS 
			DELETE FROM [SRC_France_Customs].[dbo].[E8]
			WHERE [PERIOD] IN (SELECT DISTINCT PERIOD FROM [SRC_France_Customs].[dbo].[TEMP_STEP6_EXP]);
			INSERT INTO [SRC_France_Customs].[dbo].[E8]
			SELECT *
			FROM [SRC_France_Customs].[dbo].[TEMP_STEP6_EXP];

			-- IMPORTS 
			DELETE FROM [SRC_France_Customs].[dbo].[I8]
			WHERE [PERIOD] IN (SELECT DISTINCT PERIOD FROM [SRC_France_Customs].[dbo].[TEMP_STEP6_IMP]);
			INSERT INTO [SRC_France_Customs].[dbo].[I8]
			SELECT *
			FROM [SRC_France_Customs].[dbo].[TEMP_STEP6_IMP];
			''')
			cursor.commit()
			# drop tables
			cursor.execute('''
			-- DROP TABLE [SRC_France_Customs].[dbo].[TEMP];
			DROP TABLE [SRC_France_Customs].[dbo].[TEMP_STEP_NEW];
			DROP TABLE [SRC_France_Customs].[dbo].[TEMP_STEP_OLD];
			DROP TABLE [SRC_France_Customs].[dbo].[TEMP_STEP1];
			DROP TABLE [SRC_France_Customs].[dbo].[TEMP_STEP2];
			DROP TABLE [SRC_France_Customs].[dbo].[TEMP_STEP3];
			DROP TABLE [SRC_France_Customs].[dbo].[TEMP_STEP4];
			DROP TABLE [SRC_France_Customs].[dbo].[TEMP_STEP5];
			DROP TABLE [SRC_France_Customs].[dbo].[TEMP_STEP6_EXP];
			DROP TABLE [SRC_France_Customs].[dbo].[TEMP_STEP6_IMP];
			''')
			cursor.commit()
			
			# Unit Checks
			try:
				conn, cursor = self.autoUnitCheck(conn, cursor)
			except Exception as e:
				self.StatusEmail('autoUnitCheck() Error', e, traceback.format_exc())
			insrtCondition = True
			shutil.copy(f'{self.logPath}\\RUNTIME.log', f'{self.logPath}\\{self.nextPeriodOfTDMdata}.log')
			self.logger.info('\t\tData Verified.')
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
		self.logger.info(f'Launching {self.country} ETL')

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
				self.logger.exception('dataCheck() Error')
				self.StatusEmail('dataCheck() Error', e, traceback.format_exc())
				sys.exit('dataCheck() Error')

		self.logger.info(f'{self.country} ETL Complete')


if __name__ == '__main__':

	fr = ETL('France Customs')
	fr.run()
