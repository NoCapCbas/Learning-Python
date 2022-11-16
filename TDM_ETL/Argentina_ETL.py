import paramiko
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
import random
from time import sleep
from bs4 import BeautifulSoup as bs
from pywintypes import com_error
from pathlib import Path
from pandasql import sqldf
import win32com.client as win32
import mariadb
import zipfile
from io import BytesIO
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
from selenium.webdriver.support.ui import Select
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

		self.lastPeriodOfTDMdata = list(pd.read_sql_query(Q, conn)['StopYM'])[0]
		# self.lastPeriodOfTDMdata = '202206' # hard set for testing
		conn.close()

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
		# self.nextPeriodOfTDMdata = '202207' # hard set for testing
		
		self.current_year = datetime.now().year

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
		else:
			shutil.rmtree(self.downloadPath)
			sleep(3)
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
		self.logger.info('Argentina_ETL')
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
		# open browser
		chrome_options = webdriver.ChromeOptions()
		prefs = {'download.default_directory': self.downloadPath}
		chrome_options.add_argument('--headless')
		chrome_options.add_experimental_option('prefs', prefs)
		executable_path = ChromeDriverManager().install()
		driver = webdriver.Chrome(executable_path=executable_path, chrome_options=chrome_options)
		driver.get('https://comex.indec.gov.ar/#/database')

		self.logger.info('\t\tGrabbing Imports...')
		# get number of files in downloadPath, should be zero
		num_of_files = len(os.listdir(self.downloadPath))
		# select import flow
		xPathImports = '//*[@id="root"]/main/div[3]/div[1]/select/option[2]'
		importL = WebDriverWait(driver, 120).until(EC.presence_of_element_located(((By.XPATH, xPathImports))))
		importL.click()
		# select year
		xPathTopYear = '//*[@id="root"]/main/div[3]/div[2]/select/option[2]'
		yearL = WebDriverWait(driver, 45).until(EC.presence_of_element_located(((By.XPATH, xPathTopYear))))
		yearL.click()
		# select monthly freq
		xPathMonthly = '//*[@id="root"]/main/div[3]/div[3]/select/option[2]'
		freqL = WebDriverWait(driver, 45).until(EC.presence_of_element_located(((By.XPATH, xPathMonthly))))
		freqL.click()
		# download file
		xPathDownload = '//*[@id="root"]/main/div[3]/div[4]/button'
		downloadL = WebDriverWait(driver, 45).until(EC.presence_of_element_located(((By.XPATH, xPathDownload))))
		downloadL.click()
		# wait for file to download
		fileDownloaded = False
		while fileDownloaded == False:
			sleep(3)
			
			if zipfile.is_zipfile(self.downloadPath + f'\\imports_{self.yearToQuery}_M.zip'):
				fileDownloaded = True
		
		self.logger.info('\t\tGrabbing Exports...')
		# get number of files in downloadPath, should be zero
		num_of_files = len(os.listdir(self.downloadPath))
		# select export flow
		xPathExports = '//*[@id="root"]/main/div[3]/div[1]/select/option[3]'
		exportL = WebDriverWait(driver, 120).until(EC.presence_of_element_located(((By.XPATH, xPathExports))))
		exportL.click()
		# select year
		xPathTopYear = '//*[@id="root"]/main/div[3]/div[2]/select/option[2]'
		yearL = WebDriverWait(driver, 45).until(EC.presence_of_element_located(((By.XPATH, xPathTopYear))))
		yearL.click()
		# select monthly freq
		xPathMonthly = '//*[@id="root"]/main/div[3]/div[3]/select/option[2]'
		freqL = WebDriverWait(driver, 45).until(EC.presence_of_element_located(((By.XPATH, xPathMonthly))))
		freqL.click()
		# download file
		xPathDownload = '//*[@id="root"]/main/div[3]/div[4]/button'
		downloadL = WebDriverWait(driver, 45).until(EC.presence_of_element_located(((By.XPATH, xPathDownload))))
		downloadL.click()
		# wait for file to download
		fileDownloaded = False
		while fileDownloaded == False:
			sleep(3)
			
			if zipfile.is_zipfile(self.downloadPath + f'\\exports_{self.yearToQuery}_M.zip'):
				fileDownloaded = True

		# check if data is new
		for file in os.listdir(self.downloadPath):
			if 'imports' in file:
				file_to_read = 'impom22.csv'
			if 'exports' in file:
				file_to_read = 'exponm22.csv'

			# read zip folder
			zf = zipfile.ZipFile(self.downloadPath + f'\\{file}', 'r')
			# read data file in zipfolder
			# print(zf.namelist())
			fileObj = zf.read(file_to_read)
			dataInMem = BytesIO(fileObj)
			tempDF = pd.read_csv(dataInMem, encoding='latin-1', sep=';')
			dataLatestMonAsInt = max(tempDF[tempDF.columns[1]].unique())
			zf.close()

			if int(self.nextMon) == dataLatestMonAsInt:
				dataToLoad = True
			else:
				dataToLoad = False
				break
		#
		if dataToLoad == True:
			self.logger.info(f'\t\t\t*New Data Available*')
			self.StatusEmail('New Data', f'New {self.country} Data', '')
		else:
			self.logger.info(f'\t\t\tNo New Data Available')
			self.logger.info(f'\t\t\t\tLooking for {self.nextPeriodOfTDMdata}')
			self.logger.info(f'\t\t\t\tData Latest Month {dataLatestMonAsInt} does not match {int(self.nextMon)}')
			self.logger.info(f'\tData Extraction Terminated.')
		
		self.logger.info('\tData Extraction Complete.')
		return dataToLoad

	def loadData(self):
		self.logger.info('\tLoading Data...')

		# Create Connection to [SRC_Argentina]
		conn = pyodbc.connect(
			driver = '{ODBC Driver 17 for SQL Server}', 
			server = 'SF-PROCESS',
			database = 'SRC_Argentina',
			uid = 'sa',
			pwd = 'Harpua88',
		)
		cursor = conn.cursor()

		for file in os.listdir(self.downloadPath):
			
			if 'imports' in file:
				file_to_read = 'impom22.csv'
				
			if 'exports' in file:
				file_to_read = 'exponm22.csv'

			self.logger.info(f'\t\tReading {file} {file_to_read}...')
		
			# read zip folder
			zf = zipfile.ZipFile(self.downloadPath + f'\\{file}', 'r')
			# read data file in zipfolder
			# print(zf.namelist())
			fileObj = zf.read(file_to_read)
			dataInMem = BytesIO(fileObj)
			MASTER = pd.read_csv(dataInMem, encoding='latin-1', sep=';', dtype=str)
			zf.close()
			MASTER = MASTER.astype(str)
			if 'imports' in file:
				table = 'IMP_DETL'
				MASTER = MASTER.filter([MASTER.columns[0],MASTER.columns[1],MASTER.columns[2],MASTER.columns[3],MASTER.columns[4],MASTER.columns[-1]], axis=1)
			if 'exports' in file:
				table = 'EXP_DETL'
				MASTER = MASTER.filter([MASTER.columns[0],MASTER.columns[1],MASTER.columns[2],MASTER.columns[3],MASTER.columns[4],MASTER.columns[-1]], axis=1)
			# print(MASTER)
			# print(MASTER.columns)
			
			# Load MASTER to DB
			self.logger.info(f'\t\t\tDROPPING TABLE SRC_Argentina.dbo.{table}')
			cursor.execute(f"IF OBJECT_ID('SRC_Argentina.dbo.{table}') IS NOT NULL TRUNCATE TABLE [SRC_Argentina].[dbo].[{table}];")
			cursor.commit()

			# Loop through list of split df and insert into table
			self.logger.info(f'\t\t\tINSERTING {len(MASTER)} ROWS INTO SRC_Argentina.dbo.{table}')
			self.logger.info('')
			insert_to_temp_table = f'INSERT INTO SRC_Argentina.dbo.{table} VALUES(?,?,?,?,?,?)'
			cursor.fast_executemany = True
			
			# INSERT INTO TEMP
			masterValues = MASTER.values.tolist()
			rowStepper = 1000000
			# Inserting 1000000 rows of data at a time to prevent error
			for rowI in range(0,len(masterValues), rowStepper):
				cursor.executemany(insert_to_temp_table, masterValues[rowI:rowI+rowStepper])
				cursor.commit()

			# move excel file to archive
			shutil.move(f'{self.downloadPath}\\{file}', f'{self.archivePath}\\{file}')

		cursor.close()
		conn.close()
		self.logger.info('\tData Loaded.')

	def processData(self):
		self.logger.info('\tProcessing Data...')
		# Grab SQL processing script
		sqlScriptPath = "Y:\\_PyScripts\\Damon\\Argentina\\Argentina_pyscript.sql"
		sql = self.getSql(sqlScriptPath)
		# Create Connection
		conn = pyodbc.connect(
			driver = '{ODBC Driver 17 for SQL Server}', 
			server = 'SF-PROCESS',
			database = 'SRC_Argentina', 
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
				FROM [SRC_Argentina].[DBO].[RUNNING-STATUS];
				''').fetchone()
			
			except:
				continue
			STATUS = query[0]
			if STATUS == 0:
				self.logger.info('\t\t\tSQL Script Executed.')
				running = False 
				cursor.execute('''
				IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[SRC_Argentina].[dbo].[RUNNING-STATUS]') AND type in (N'U'))
				DROP TABLE [SRC_Argentina].[DBO].[RUNNING-STATUS];
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
			database = 'SRC_Argentina', 
			uid = 'sa', 
			pwd = 'Harpua88'
		)
		cursor = conn.cursor()

		# Generic Checks
		self.logger.info('\t\tGeneric Checks...')

		PRE_FINAL_TABLE_EXP = '[SRC_Argentina].[dbo].[EXP_STEP3]'
		PRE_FINAL_TABLE_IMP = '[SRC_Argentina].[dbo].[IMP_STEP3]'

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
		for file in os.listdir(self.archivePath):
			if 'imports' in file:
				file_to_read = 'itotm22.csv'
			if 'exports' in file:
				file_to_read = 'etotnm22.csv'
			self.logger.info(f'\t\t\tReading {file} {file_to_read}...')
			# read zip folder
			zf = zipfile.ZipFile(self.archivePath + f'\\{file}', 'r')
			# read data file in zipfolder
			# print(zf.namelist())
			fileObj = zf.read(file_to_read)
			dataInMem = BytesIO(fileObj)
			MASTER = pd.read_csv(dataInMem, encoding='latin-1', sep=';', dtype=str)
			zf.close()
			MASTER = MASTER.astype(str)
			if 'imports' in file:
				table = 'IMP_DETL'
				MASTER_IMP = MASTER.filter([MASTER.columns[0],MASTER.columns[-1],MASTER.columns[4]], axis=1)
				MASTER_IMP = MASTER_IMP.rename(columns={
					MASTER_IMP.columns[0]:'SOURCE FILE',
					MASTER_IMP.columns[1]:'SOURCE VALUE',
					MASTER_IMP.columns[2]:'SOURCE QTY',
				})
			if 'exports' in file:
				table = 'EXP_DETL'
				MASTER_EXP = MASTER.filter([MASTER.columns[0],MASTER.columns[-1],MASTER.columns[4]], axis=1)
				MASTER_EXP = MASTER_EXP.rename(columns={
					MASTER_EXP.columns[0]:'SOURCE FILE',
					MASTER_EXP.columns[1]:'SOURCE VALUE',
					MASTER_EXP.columns[2]:'SOURCE QTY',
				})
			# print(MASTER)
			# print(MASTER.columns)

		EXP_TOTALS = pd.read_sql_query(f'''
		SELECT CAST(SUM([VALUE]) AS VARCHAR) AS [PROCESSED TOTAL]
		,CAST(SUM([QTY1]) AS VARCHAR) AS [PROCESSED QTY1]
		FROM {PRE_FINAL_TABLE_EXP}
		''',conn)
		IMP_TOTALS = pd.read_sql_query(f'''
		SELECT CAST(SUM([VALUE]) AS VARCHAR) AS [PROCESSED TOTAL]
		,CAST(SUM([QTY1]) AS VARCHAR) AS [PROCESSED QTY1]
		FROM {PRE_FINAL_TABLE_IMP}
		''',conn)

		# Check if newest period exists within final tables before inserting, want a table of length zero 
		check_E8 = pd.read_sql_query(f'''
		SELECT DISTINCT a.PERIOD, b.PERIOD
		FROM {PRE_FINAL_TABLE_EXP} a
		LEFT JOIN [SRC_Argentina].[dbo].[E8] b
		ON a.PERIOD = b.PERIOD
		WHERE b.PERIOD IS NOT NULL
		AND a.PERIOD = '{self.nextPeriodOfTDMdata}';
		''', conn)
		# print(len(check_E8))
		check_I8 = pd.read_sql_query(f'''
		SELECT DISTINCT a.PERIOD, b.PERIOD
		FROM {PRE_FINAL_TABLE_IMP} a
		LEFT JOIN [SRC_Argentina].[dbo].[I8] b
		ON a.PERIOD = b.PERIOD
		WHERE b.PERIOD IS NOT NULL
		AND a.PERIOD = '{self.nextPeriodOfTDMdata}';
		''', conn)
		# print(len(check_I8))

		insrtCondition = False
		if list(genericChecks.all())[0] and len(check_E8) == 0 and len(check_I8) == 0:
			# insert into archive 
			cursor.execute('''
			--Delete from the Archive Cumulative Monthly Data for the Year
			DELETE FROM [SRC_Argentina].[dbo].[EXP_DETL_ARCHIVE]
			Where [Anio] in (select distinct [Anio] from [SRC_Argentina].[dbo].[EXP_DETL])

			--Insert New Cumulative Monthly Data into the Archive
			INSERT INTO [SRC_Argentina].[dbo].[EXP_DETL_ARCHIVE] 
			SELECT *
			FROM [EXP_DETL_Step0]

			--Delete from the Archive Cumulative Monthly Data for the Year
			DELETE FROM [SRC_Argentina].[dbo].[IMP_DETL_ARCHIVE]
			Where [Anio] in (select distinct [Anio] from [SRC_Argentina].[dbo].[IMP_DETL])

			--Insert New Cumulative Monthly Data into the Archive
			INSERT INTO [SRC_Argentina].[dbo].[IMP_DETL_ARCHIVE] 
			SELECT *
			FROM [IMP_DETL_Step0]
			''') 
			cursor.commit()

			# insert into final tables
			cursor.execute('''
			--Delete revised month in E8 before inserting data to E8
			DELETE FROM [SRC_Argentina].[dbo].[E8]
			WHERE [YR] in (select distinct YR from [SRC_Argentina].[dbo].[EXP_STEP3]);
			 
			--Final Insert into E8 
			INSERT INTO [SRC_Argentina].[dbo].[E8]
			SELECT *
			FROM [SRC_Argentina].[dbo].[EXP_STEP3];

			--Delete revised month in I8 before inserting data to I8
			DELETE FROM [SRC_Argentina].[dbo].[I8]
			WHERE [YR] in (select distinct YR from [SRC_Argentina].[dbo].[IMP_STEP3]);
			 
			--Final Insert into I8 
			INSERT INTO [SRC_Argentina].[dbo].[I8]
			SELECT *
			FROM [SRC_Argentina].[dbo].[IMP_STEP3];
			''')
			cursor.commit()

			# drop tables
			cursor.execute('''
			DROP TABLE [SRC_Argentina].[dbo].[EXP_DETL_Step0];
			DROP TABLE [SRC_Argentina].[dbo].[EXP_STEP1];
			DROP TABLE [SRC_Argentina].[dbo].[EXP_STEP2];
			DROP TABLE [SRC_Argentina].[dbo].[EXP_STEP2_a];
			DROP TABLE [SRC_Argentina].[dbo].[EXP_STEP3];
			DROP TABLE [SRC_Argentina].[dbo].[IMP_DETL_Step0];
			DROP TABLE [SRC_Argentina].[dbo].[IMP_STEP1];
			DROP TABLE [SRC_Argentina].[dbo].[IMP_STEP2];
			DROP TABLE [SRC_Argentina].[dbo].[IMP_STEP2_a];
			DROP TABLE [SRC_Argentina].[dbo].[IMP_STEP3];
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
					{MASTER_EXP.to_html(index=False)}
					<br>
					{EXP_TOTALS.to_html(index=False)}
				<br>
				<hr>
				<h3>Totals: IMPORT</h3>
					{MASTER_IMP.to_html(index=False)}
					<br>
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

	ar = ETL('Argentina')
	ar.run()