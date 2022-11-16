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
from selenium.webdriver.chrome.options import Options
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
		#self.nextPeriodOfTDMdata = '202206' # hard set for testing
		
		self.current_year = datetime.now().year
		# URLS
		self.data_source_url = 'https://www.bps.go.id/exim/'
		# PATHS
		self.logPath = f"Y:\\_PyScripts\\Damon\\{self.country}\\Log"
		self.downloadPath = f"Y:\\_PyScripts\\Damon\\{self.country}\\Downloads"
		self.archivePath = f'Y:\\{self.country}\\Archive\\{self.nextPeriodOfTDMdata}'
		self.totalsPath = f'Y:\\{self.country}\\Check Totals\\{self.nextPeriodOfTDMdata}'
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
		if not os.path.exists(self.totalsPath): 
			os.makedirs(self.totalsPath)
		
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
		self.logger.info('Indonesia_ETL')
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
			FROM [SRC_Indonesia].[dbo].[{k}8]
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

	def checkLatestPeriodProcessed(self):
		# Create Connection
		conn = pyodbc.connect(
			driver = '{ODBC Driver 17 for SQL Server}', 
			server = 'SF-PROCESS', 
			database = 'SRC_Indonesia', 
			uid = 'sa', 
			pwd = 'Harpua88')

		E8_StopYM = pd.read_sql_query(F'''
		SELECT DISTINCT PERIOD
		FROM [SRC_Indonesia].[dbo].[E8]
		WHERE PERIOD = {self.nextPeriodOfTDMdata}
		''',conn)

		I8_StopYM = pd.read_sql_query(F'''
		SELECT DISTINCT PERIOD
		FROM [SRC_Indonesia].[dbo].[I8]
		WHERE PERIOD = {self.nextPeriodOfTDMdata}
		''',conn)

		if len(E8_StopYM) == 0 and len(I8_StopYM) == 0:
			return True
		else:
			return False

	def extractData(self):
		# checks DB1 if next period exists already but has not been published
		dataToLoad = self.checkLatestPeriodProcessed()
		if dataToLoad == False:
			return dataToLoad

		self.logger.info('\tExtracting Data...')
		gettingRequest = True
		while gettingRequest:
			try:
				# grab cookies
				r = requests.get('https://www.bps.go.id/exim/')
				cookies = r.cookies.get_dict()
				cookie_string = "; ".join([str(x)+"="+str(y) for x,y in cookies.items()])
				# print(cookie_string)
				gettingRequest = False
			except:
				self.logger.info('\t\tRequest Failed.')


		soup = bs(r.text)
		# print(soup)
		# grab form submission token
		tokens = soup.find_all('input', {'name':'YII_CSRF_TOKEN'})
		# print(tokens)
		token = tokens[0]['value']
		# print(token)
		
		# might be useful to compare totals, spit out a totals file in the archive
		# then use in the dataChecks function to compare
		# urlTOTAL = "https://www.bps.go.id/exim/getDataSummary.html"
		# payload = f"YII_CSRF_TOKEN={token}&bulan={self.nextMon}&tahun={self.yearToQuery}&filter=filtered"
		# r = requests.request('POST', urlTOTAL, data=payload)
		# total_soup = bs(r.text)
		# total_html = total_soup.find('table', {'id':'summaryexim'})
		# print(total_html)
		# sys.exit()
		
		col2 = soup.find_all('div', {'id':'column2'})[0]
		# print(col2)
		
		# grab latest date published by source
		latestPublishedDataMonth, latestPublishedDataYear = col2.find_all('b')[0].text.split(' ')
		# print(latestPublishedDataMonth)
		
		month_conc = {
			'Januari':'01',
			'Februari':'02',
			'Maret':'03',
			'April':'04',
			'Mei':'05',
			'Juni':'06',
			'Juli':'07',
			'Agustus':'08',
			'September':'09',
			'Oktober':'10',
			'November':'11',
			'Desember':'12',
		}
		# check if data is new 
		if self.nextPeriodOfTDMdata == latestPublishedDataYear + month_conc[latestPublishedDataMonth]:
			self.logger.info(f'\t\t{self.nextPeriodOfTDMdata} = {latestPublishedDataYear + month_conc[latestPublishedDataMonth]}')
			self.logger.info('\t\t*New Data Available*')
			self.StatusEmail('New Data', f'New {self.country} Data', '')
			dataToLoad = True
		else:
			self.logger.info(f'\t\t{self.nextPeriodOfTDMdata} != {latestPublishedDataYear + month_conc[latestPublishedDataMonth]}')
			self.logger.info('\t\tNo New Data Available.')
			self.logger.info(f'\tData Extraction Terminated.')
			dataToLoad = False
			return dataToLoad

		# grab list of countries to scrape 
		cty_selection = soup.find_all('select', {'id':'negara_ctry'})[0]
		cty_tags = cty_selection.find_all('option')

		headers = {
		'Accept': 'application/json, text/javascript, */*; q=0.01',
		'Accept-Language': 'en-US,en;q=0.9',
		'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
		'Cookie': cookie_string,
		'Origin': 'https://www.bps.go.id',
		'Referer': 'https://www.bps.go.id/exim',
		}

		url = self.data_source_url + 'processnew.html'
		# begin making requests to grab data
		for flow in ['EXP', 'IMP']:
			if flow == 'EXP':
				flowCode = 1
			if flow == 'IMP':
				flowCode = 2
			for cty_tag in cty_tags:
				sleep(random.uniform(3,15))
				cty_iso = cty_tag['value']
				
				payload = f"menurut%5B%5D=2&kelompokhs=&sumber={flowCode}&kodehs=&port=&ctry%5B%5D={cty_iso}&bulan%5B%5D={self.nextMon}&tahun%5B%5D={self.yearToQuery}&YII_CSRF_TOKEN={token}"
				# payload = f"menurut%5B%5D=2&kelompokhs=&sumber={flowCode}&kodehs=&port=&ctry%5B%5D={cty_iso}&bulan%5B%5D={'06'}&tahun%5B%5D={self.yearToQuery}&YII_CSRF_TOKEN={token}" # hard set for testing
				self.logger.info(f'\t\t\tDownloading {payload}...')
				r = requests.post(url, headers=headers, data=payload)
				# print(r.text)
				tempDF = pd.read_json(r.text)
				tempDF.to_csv(f'{self.downloadPath}\\{flow}{self.nextPeriodOfTDMdata}_{cty_iso}.csv', index=False)
				# sys.exit()

		# haven't test might break
		# grab source totals 
		options = Options()
		options.headless = True
		driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
		driver.get('https://www.bps.go.id/exim/')
		
		select = Select(driver.find_element_by_id('filterthn'))
		select.select_by_visible_text(self.yearToQuery)
		l = WebDriverWait(driver, 120).until(EC.element_to_be_clickable(((By.ID, 'summaryexim')))).get_attribute('outerHTML')
		tempDF = pd.read_html(l.replace(u'\xa0', ''))[0]
		# print(tempDF)
		
		PERIODS = [f'{self.yearToQuery}{i}' for i in range(1,13)]
		valueEXP = [i.replace('\xa0', '').replace(',', '.') for i in tuple(tempDF['Nilai Ekspor (US $)'])]
		qtyEXP = [i.replace('\xa0', '').replace(',', '.') for i in tuple(tempDF['Berat Ekspor (KG)'])]
		valueIMP = [i.replace('\xa0', '').replace(',', '.') for i in tuple(tempDF['Nilai Impor (US $)'])]
		qtyIMP = [i.replace('\xa0', '').replace(',', '.') for i in tuple(tempDF['Berat Impor (KG)'])]

		tempDF = pd.DataFrame({
			'PERIOD':PERIODS,
			'VALUE EXPORTS':valueEXP,
			'QTY EXPORTS':qtyEXP,
			'VALUE IMPORTS':valueIMP,
			'QTY IMPORTS':qtyIMP,
		})

		tempDF.to_csv(f'{self.totalsPath}\\{self.nextPeriodOfTDMdata}_Totals_pyscript.csv', index=False)

		driver.quit()

		self.logger.info('\tData Extraction Complete.')
		return dataToLoad

	def loadData(self):
		self.logger.info('\tLoading Data...')
		MASTER = pd.DataFrame()
		# create connection to [SRC_Indonesia].[dbo].[TEMP]
		conn = pyodbc.connect(
			driver = '{ODBC Driver 17 for SQL Server}', 
			server = 'SF-PROCESS', 
			database = 'SRC_Indonesia', 
			uid = 'sa', 
			pwd = 'Harpua88')
		cursor = conn.cursor()

		# loop through files 
		for file in os.listdir(self.downloadPath):
			self.logger.info(f'\t\tReading {file}...')
			try:
				tempDF = pd.read_csv(f'{self.downloadPath}\\{file}')
			except:
				tempDF = pd.DataFrame()
			
			# Add tempDF to MASTER
			if len(tempDF) > 0:
				tempDF.insert(0, 'FILENAME', f'{self.archivePath}\\{file}')
				MASTER = pd.concat([MASTER, tempDF], ignore_index=True)
			else:
				pass

			# moves file to archive once read
			shutil.move(f'{self.downloadPath}\\{file}', f'{self.archivePath}\\{file}')
		self.logger.info('\t\tAll Data Read.')

		# print(MASTER)
		MASTER = MASTER.astype(str)
		# print(MASTER.columns)
		MASTER = MASTER.rename(columns={
			MASTER.columns[1]:'VALUE', 
			MASTER.columns[2]:'NETWEIGHT', 
			MASTER.columns[3]:'HS CODE', 
			MASTER.columns[4]:'PELABUHAN', 
			MASTER.columns[5]:'PTN_CTY', 
			MASTER.columns[6]:'YEAR', 
			MASTER.columns[7]:'MONTH', 
		})
		

		# Load MASTER to DB
		self.logger.info('\t\t\tDROPPING TABLE SRC_Indonesia.dbo.TEMP')
		cursor.execute(f"IF OBJECT_ID('SRC_Indonesia.dbo.TEMP') IS NOT NULL DROP TABLE [SRC_Indonesia].[dbo].[TEMP];")
		cursor.commit()

		# CREATE TABLE
		create_statements_cols = ''
		for col in MASTER.columns:
			if MASTER.columns[-1] == col:
				create_statements_cols = create_statements_cols + f'[{col}] [varchar](max) NULL'
			else: 
				create_statements_cols = create_statements_cols + f'[{col}] [varchar](max) NULL,'

		self.logger.info(f'\t\t\tCREATING TABLE SRC_Indonesia.dbo.TEMP')
		cursor.execute(f"""
		CREATE TABLE SRC_Indonesia.dbo.TEMP(
			{create_statements_cols}
		)
		""")
		cursor.commit()

		# Loop through list of split df and insert into table
		self.logger.info(f'\t\t\tINSERTING {len(MASTER)} ROWS INTO SRC_Indonesia.dbo.TEMP')
		self.logger.info('')
		insert_to_temp_table = f'INSERT INTO SRC_Indonesia.dbo.TEMP VALUES(?,?,?,?,?,?,?,?)'
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
		sqlScriptPath = "Y:\\_PyScripts\\Damon\\Indonesia\\Indonesia_pyscript.sql"
		sql = self.getSql(sqlScriptPath)
		# Create Connection
		conn = pyodbc.connect(
			driver = '{ODBC Driver 17 for SQL Server}', 
			server = 'SF-PROCESS',
			database = 'SRC_Indonesia', 
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
				FROM [SRC_Indonesia].[DBO].[RUNNING-STATUS];
				''').fetchone()
			
			except:
				continue
			STATUS = query[0]
			if STATUS == 0:
				self.logger.info('\t\t\tSQL Script Executed.')
				running = False 
				cursor.execute('''
				IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[SRC_Indonesia].[dbo].[RUNNING-STATUS]') AND type in (N'U'))
				DROP TABLE [SRC_Indonesia].[DBO].[RUNNING-STATUS];
				''')
				cursor.commit()
			else:
				running = True

		cursor.close()
		conn.close()
		self.logger.info('\tData Processed.')

	def dataChecks(self):
		self.logger.info('\tPerforming Checks...')

		# Create Connnection
		conn = pyodbc.connect(
			driver = '{ODBC Driver 17 for SQL Server}', 
			server = 'SF-PROCESS', 
			database = 'SRC_Indonesia', 
			uid = 'sa', 
			pwd = 'Harpua88'
		)
		
		cursor = conn.cursor()

		# Generic Checks
		self.logger.info('\t\tGeneric Checks...')

		PRE_FINAL_TABLE_EXP = '[SRC_Indonesia].[dbo].[SRC_EXP_STEP5]'
		PRE_FINAL_TABLE_IMP = '[SRC_Indonesia].[dbo].[SRC_IMP_STEP5]'

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
			'E8 COMMODITY_LEN IS 10 DIGITS': [False], 
			'I8 COMMODITY_LEN IS 10 DIGITS': [False],
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

		if dfE8['COMMODITY'].astype(str).map(len).nunique() == 1 and dfE8['COMMODITY'].astype(str).map(len).unique()[0] == 10:
			failedChecks['E8 COMMODITY_LEN IS 10 DIGITS'][0] = True

		if dfI8['COMMODITY'].astype(str).map(len).nunique() == 1 and dfI8['COMMODITY'].astype(str).map(len).unique()[0] == 10:
			failedChecks['I8 COMMODITY_LEN IS 10 DIGITS'][0] = True

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
		EXP_TOTALS = pd.read_sql_query(f'''
		SELECT [PERIOD]
		, cast(sum(cast([VALUE] as decimal(38,8))) as varchar) as [PROCESSED VALUE]
		, cast(sum(cast([QTY1] as decimal(38,8))) as varchar) as [PROCESSED QTY1]
		, cast(sum(cast([QTY2] as decimal(38,8))) as varchar) as [PROCESSED QTY2]
		FROM {PRE_FINAL_TABLE_EXP}
		GROUP BY PERIOD
		ORDER BY 1
		''',conn)
		IMP_TOTALS = pd.read_sql_query(f'''
		SELECT [PERIOD]
		, cast(sum(cast([VALUE] as decimal(38,8))) as varchar) as [PROCESSED VALUE]
		, cast(sum(cast([QTY1] as decimal(38,8))) as varchar) as [PROCESSED QTY1]
		, cast(sum(cast([QTY2] as decimal(38,8))) as varchar) as [PROCESSED QTY2]
		FROM {PRE_FINAL_TABLE_IMP}
		GROUP BY PERIOD
		ORDER BY 1
		''',conn)

		insrtCondition = False
		if list(genericChecks.all())[0]:
			# insert archive table
			cursor.execute('''
			INSERT INTO [SRC_Indonesia].[dbo].[SRC_EXP_IMP_ARCHIVE]
			SELECT [FILENAME]
			,[FLOW]
			,[HSCODE]
			,[DESCRIPTION]
			,[YEAR]
			,[MONTH]
			,[Partner_CTY]
			,[value_USD]
			,[Quantity_KG]
			FROM [SRC_Indonesia].[dbo].[SRC_EXP_IMP];
			''')
			# insert into final tables
			cursor.execute('''
			-- EXPORTS 
			INSERT INTO [SRC_Indonesia].[dbo].[E8]
  			SELECT 	
				cast([CTY_RPT] AS VARCHAR(3)) AS [CTY_RPT]
				,cast([CTY_PTN] AS VARCHAR(7)) AS [CTY_PTN]
				,cast([COMMODITY] AS VARCHAR(12)) AS [COMMODITY]
				,cast([PERIOD] AS VARCHAR(6)) AS [PERIOD]
				,cast([YR] AS INT) AS [YR]
				,cast([MON] AS TINYINT) AS [MON]
				,cast([VALUE] AS DECIMAL(38,8)) AS [VALUE]
				,cast([UNIT1] AS VARCHAR(3)) AS [UNIT1]
				,cast([QTY1] AS DECIMAL(18,2)) AS [QTY1]
				,cast([UNIT1] AS VARCHAR(3)) AS [UNIT2]
				,cast([QTY1] AS DECIMAL(18,2)) AS [QTY2]
			FROM [SRC_Indonesia].[dbo].[SRC_EXP_IMP_STEP4.5]
			WHERE FLOW = 'EXP';

			-- IMPORTS
			INSERT INTO [SRC_Indonesia].[dbo].[I8]
  			SELECT 	
				cast([CTY_RPT] AS VARCHAR(3)) AS [CTY_RPT]
				,cast([CTY_PTN] AS VARCHAR(7)) AS [CTY_PTN]
				,cast([COMMODITY] AS VARCHAR(12)) AS [COMMODITY]
				,cast([PERIOD] AS VARCHAR(6)) AS [PERIOD]
				,cast([YR] AS INT) AS [YR]
				,cast([MON] AS TINYINT) AS [MON]
				,cast([VALUE] AS DECIMAL(38,8)) AS [VALUE]
				,cast([UNIT1] AS VARCHAR(3)) AS [UNIT1]
				,cast([QTY1] AS DECIMAL(18,2)) AS [QTY1]
				,cast([UNIT1] AS VARCHAR(3)) AS [UNIT2]
				,cast([QTY1] AS DECIMAL(18,2)) AS [QTY2]
			FROM [SRC_Indonesia].[dbo].[SRC_EXP_IMP_STEP4.5]
			WHERE FLOW = 'IMP';

			-- EXPORTS ADF
			INSERT INTO [SRC_Indonesia].[dbo].[E8_ADF]
  			SELECT 	
				cast([CTY_RPT] AS VARCHAR(3)) AS [CTY_RPT]
				,cast([CTY_PTN] AS VARCHAR(7)) AS [CTY_PTN]
				,cast([COMMODITY] AS VARCHAR(12)) AS [COMMODITY]
				,cast([PERIOD] AS VARCHAR(6)) AS [PERIOD]
				,cast([YR] AS INT) AS [YR]
				,cast([MON] AS TINYINT) AS [MON]
				,cast([VALUE] AS DECIMAL(38,8)) AS [VALUE]
				,cast([UNIT1] AS VARCHAR(3)) AS [UNIT1]
				,cast([QTY1] AS DECIMAL(18,2)) AS [QTY1]
				,cast([UNIT2] AS VARCHAR(3)) AS [UNIT2]
				,cast([QTY2] AS DECIMAL(18,2)) AS [QTY2]
				,[PORT]
			FROM [SRC_Indonesia].[dbo].[SRC_EXP_STEP5];

			-- IMPORTS ADF
			INSERT INTO [SRC_Indonesia].[dbo].[I8_ADF]
  			SELECT 	
				cast([CTY_RPT] AS VARCHAR(3)) AS [CTY_RPT]
				,cast([CTY_PTN] AS VARCHAR(7)) AS [CTY_PTN]
				,cast([COMMODITY] AS VARCHAR(12)) AS [COMMODITY]
				,cast([PERIOD] AS VARCHAR(6)) AS [PERIOD]
				,cast([YR] AS INT) AS [YR]
				,cast([MON] AS TINYINT) AS [MON]
				,cast([VALUE] AS DECIMAL(38,8)) AS [VALUE]
				,cast([UNIT1] AS VARCHAR(3)) AS [UNIT1]
				,cast([QTY1] AS DECIMAL(18,2)) AS [QTY1]
				,cast([UNIT2] AS VARCHAR(3)) AS [UNIT2]
				,cast([QTY2] AS DECIMAL(18,2)) AS [QTY2]
				,[PORT]
			FROM [SRC_Indonesia].[dbo].[SRC_IMP_STEP5];
			''')
			cursor.commit()
			# drop tables
			cursor.execute('''
			--DROP TABLE [SRC_Indonesia].[dbo].[TEMP];
			DROP TABLE [SRC_Indonesia].[dbo].[SRC_EXP_IMP_STEP1];
			DROP TABLE [SRC_Indonesia].[dbo].[SRC_EXP_IMP_STEP2];
			DROP TABLE [SRC_Indonesia].[dbo].[SRC_EXP_IMP_STEP3];
			DROP TABLE [SRC_Indonesia].[dbo].[SRC_EXP_IMP_STEP4];
			DROP TABLE [SRC_Indonesia].[dbo].[SRC_EXP_IMP_STEP4.5];
			DROP TABLE [SRC_Indonesia].[dbo].[SRC_EXP_STEP5];
			DROP TABLE [SRC_Indonesia].[dbo].[SRC_IMP_STEP5];
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

		# Grab Source Totals 
		SRC_TOTALS = pd.read_csv(f'{self.totalsPath}\\{self.nextPeriodOfTDMdata}_Totals_pyscript.csv')

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

				<hr>
				<h3>SOURCE TOTALS</h3>
					{SRC_TOTALS.to_html(index=False)}
				
				
				
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
		self.logger.info(f'{self.country} ETL Finished.')


if __name__ == '__main__':

	id = ETL('Indonesia')
	id.run()
