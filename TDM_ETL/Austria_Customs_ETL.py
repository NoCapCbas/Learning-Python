import paramiko
import math
import pyodbc 
import pandas as pd
import logging
from datetime import datetime
import shutil
import traceback
import sys
import os
from time import sleep
from pywintypes import com_error
from pathlib import Path
import win32com.client as win32
import mariadb
import zipfile
from io import BytesIO
# email import
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
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
		WHERE Declarant = '{self.country.replace('_', ' ')}'"""

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
		# print(self.nextPeriodOfTDMdata)
		
		
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
		self.logger.info('Austria_Customs_ETL')
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
	
	def checkLatestPeriodProcessed(self):
		# Create Connection
		conn = pyodbc.connect(
			driver = '{ODBC Driver 17 for SQL Server}', 
			server = 'SF-PROCESS', 
			database = 'SRC_Austria_Customs', 
			uid = 'sa', 
			pwd = 'Harpua88')

		E8_StopYM = pd.read_sql_query(F'''
		SELECT DISTINCT PERIOD
		FROM [SRC_Austria_Customs].[dbo].[E8]
		WHERE PERIOD = {self.nextPeriodOfTDMdata}
		''',conn)

		I8_StopYM = pd.read_sql_query(F'''
		SELECT DISTINCT PERIOD
		FROM [SRC_Austria_Customs].[dbo].[I8]
		WHERE PERIOD = {self.nextPeriodOfTDMdata}
		''',conn)

		if len(E8_StopYM) == 0 and len(I8_StopYM) == 0:
			return True
		else:
			return False

	def extractData(self):
		def progressbar(x, y):
			''' progressbar for the pysftp
			'''
			bar_len = 60
			filled_len = math.ceil(bar_len * x / float(y))
			percents = math.ceil(100.0 * x / float(y))
			bar = '=' * filled_len + '-' * (bar_len - filled_len)
			filesize = f'{math.ceil(y/1024):,} KB' if y > 1024 else f'{y} byte'
			sys.stdout.write(f'[{bar}] {percents}% {filesize}\r')
			sys.stdout.flush()
		# checks DB1 if next period exists already but has not been published
		dataToLoad = self.checkLatestPeriodProcessed()
		if dataToLoad == False:
			return dataToLoad

		self.logger.info('\tExtracting Data...')
		# clear download folder
		shutil.rmtree(self.downloadPath)
		sleep(3)
		if not os.path.exists(self.downloadPath): 
			os.makedirs(self.downloadPath)
		
		host = 'ftp.statistik.gv.at'
		port = 22
		user = 'ahtdm'
		password = 'JbuqcYqfoUDU'

		self.logger.info(f'\t\tConnecting to host:{host}')
		
		# Open a transport
		transport = paramiko.Transport((host, port))
		# SFTP FIXES
		transport.default_window_size = paramiko.common.MAX_WINDOW_SIZE
		transport.default_max_packet_size= 200 * 1024 * 1024
		transport.packetizer.REKEY_BYTES = pow(2, 40)  # 1TB max, this is a security degradation!
		transport.packetizer.REKEY_PACKETS = pow(2, 40)  # 1TB max, this is a security degradation!
		# / SFTP FIXES
		# Auth
		transport.connect(None, user, password)
		# Go
		sftp = paramiko.SFTPClient.from_transport(transport)
		files = sftp.listdir()
		
		if len(files) == 0: 
			self.logger.info('\t\tNo New Data Available.')
			return dataToLoad
		else:
			f = files[0]

		self.logger.info(f'\t\tSFTP: {files}')
		self.logger.info(f'\t\tLooking for: {self.nextPeriodOfTDMdata}')
		if files and f.split('.')[0].split('_')[1] == self.nextMon:
			dataToLoad = False
			self.StatusEmail('New Data', f'New {self.country} Data', '')
			
			# New Data is confirmed so loop getting file until it is successful
			while dataToLoad == False:
				try:
					self.logger.info(f'\t\t\tGrabbing {f}...')
					with sftp:
						sftp.get(f, f"{self.downloadPath}\\{f}",callback=lambda x,y: progressbar(x,y))
					dataToLoad = True
				except Exception as e:
					self.logger.info(e)
					self.logger.info(f'\t\t\tRetry Grabbing {f}...')
					if os.path.exists(self.downloadPath):
						shutil.rmtree(self.downloadPath)
					if not os.path.exists(self.downloadPath): 
						os.makedirs(self.downloadPath)
					dataToLoad = False
			
		else:
			self.logger.info('\t\tNo New Data Available.')
			dataToLoad = False
			return dataToLoad

		# Close sftp server
		if sftp: sftp.close()
		if transport: transport.close()

		# Clear Maria DB 
		conn = mariadb.connect(
				user="root",
				password="tdm123",
				port=3307,
				database="Austria_Customs"
			)
		members = []
		# Get Cursor
		cur = conn.cursor()
		cur.execute("SELECT * FROM information_schema.tables WHERE table_schema='Austria_Customs'")

		# drop all current tables in austria_customs mariaDB
		tables = []
		for i in cur:
			table_name = i[2]
			tables.append(table_name)

		for table in tables:
			cur.execute(f"DROP TABLE {table}")

		# extract only necassary files from zip downloaded from sftp 
		for f in os.listdir(self.downloadPath):
			self.logger.info(f'\t\t\tOpening {f}...')
			# path to zip folder 
			zipfolderPath = f'{self.downloadPath}\\{f}'
			# creating zipfolder obj
			zipRef = zipfile.ZipFile(zipfolderPath, 'r')
			# zipRef.extractall(self.downloadPath)
			for name in zipRef.namelist():
				if 'Daten' in name:
					zfiledata = BytesIO(zipRef.read(name))
					zipDaten = zipfile.ZipFile(zfiledata)
					members=[name2 for name2 in zipDaten.namelist() if 'AHCD' in name2 and '/' in name2]
					zipDaten.extractall(self.downloadPath, members=members)
					break
		fwdSlash = '/'
		bckSlash = '\\'
		# move files to mariadb directory
		for f in members:
			shutil.copy2(self.downloadPath + f'\\{f.replace(fwdSlash, bckSlash)}', f'C:\\Program Files\\MariaDB 10.9\\data\\austria_customs\\{f.split(fwdSlash)[-1]}')
			
		cur.close()
		conn.close()
		self.logger.info('\tData Extraction Complete.')
		return True

	def loadData(self):
		# Connect to MariaDB Platform
		connMaria = mariadb.connect(
			user="root",
			password="tdm123",
			port=3307,
			database="Austria_Customs"
		)
		# Get Cursor
		cursorMaria = connMaria.cursor()

		# create connection to [SRC_Austria_Customs].[dbo].[TEMP]
		conn = pyodbc.connect(
			driver = '{ODBC Driver 17 for SQL Server}', 
			server = 'SF-PROCESS', 
			database = 'SRC_Austria_Customs', 
			uid = 'sa', 
			pwd = 'Harpua88')
		cursor = conn.cursor()

		cursorMaria.execute("SELECT * FROM information_schema.tables WHERE table_schema='Austria_Customs'")

		for i in cursorMaria:
			
			maria_table_name = i[2]
			sqlServer_table_name = f'DATEN_TEMP' + maria_table_name
			if ('countries' in maria_table_name or 
				'oddlot' in maria_table_name or
				'nomkn' in maria_table_name or 
				'typeoftransaction' in maria_table_name) :
				self.logger.info(f'\t\t\tMaria Table: {maria_table_name}')


				MASTER = pd.read_sql_query(f"SELECT * FROM {maria_table_name}", connMaria)
				num_of_col = len(MASTER.columns)

				# create param str
				param_str = ''
				for x in range(num_of_col):
					if num_of_col-1 == x:
						param_str = param_str + '?'
					else:
						param_str = param_str + '?,'
				# print(param_str)

				# Load MASTER to DB
				self.logger.info(f'\t\t\tDROPPING TABLE SRC_Austria_Customs.dbo.{sqlServer_table_name}')
				cursor.execute(f"IF OBJECT_ID('SRC_Austria_Customs.dbo.{sqlServer_table_name}') IS NOT NULL DROP TABLE [SRC_Austria_Customs].[dbo].[{sqlServer_table_name}];")
				cursor.commit()

				# CREATE TABLE
				create_statements_cols = ''
				for col in MASTER.columns:
					if MASTER.columns[-1] == col:
						create_statements_cols = create_statements_cols + f'[{col}] [varchar](max) NULL'
					else: 
						create_statements_cols = create_statements_cols + f'[{col}] [varchar](max) NULL,'

				self.logger.info(f'\t\t\tCREATING TABLE SRC_Austria_Customs.dbo.{sqlServer_table_name}')
				cursor.execute(f"""
				CREATE TABLE SRC_Austria_Customs.dbo.{sqlServer_table_name}(
					{create_statements_cols}
				)
				""")
				cursor.commit()

				# Loop through list of split df and insert into table
				self.logger.info(f'\t\t\tINSERTING {len(MASTER)} ROWS INTO SRC_Austria_Customs.dbo.{sqlServer_table_name}')
				self.logger.info('')
				insert_to_temp_table = f'INSERT INTO SRC_Austria_Customs.dbo.{sqlServer_table_name} VALUES({param_str})'
				cursor.fast_executemany = True

				# INSERT INTO TEMP
				masterValues = MASTER.values.tolist()
				rowStepper = 1000000
				# Inserting 1000000 rows of data at a time to prevent error
				for rowI in range(0,len(masterValues), rowStepper):
					cursor.executemany(insert_to_temp_table, masterValues[rowI:rowI+rowStepper])
					cursor.commit()

		cursorMaria.execute("SELECT * FROM information_schema.tables WHERE table_schema='Austria_Customs'")

		tables = []
		MASTER_DATEN = pd.DataFrame()
		for i in cursorMaria:
			maria_table_name = i[2]

			if ('kn_e' in maria_table_name or 
				'kn_a' in maria_table_name):
				self.logger.info(f'\t\t\tMaria Table: {maria_table_name}')

				tempDF = pd.read_sql_query(f"SELECT * FROM {maria_table_name}", connMaria)
				len_of_df = len(tempDF)
				# we do not want the empty tables so this will skip them 
				if len_of_df == 0:
					continue
				else: 
					tables.append(int(maria_table_name.split('_')[-1]))
					tempDF.insert(0, 'TABLENAME', maria_table_name)
					MASTER_DATEN = pd.concat([MASTER_DATEN, tempDF], ignore_index=True)

		periodTableCONC = {}
		tables = list(set(tables))
		tables.sort(reverse=True)
		period = self.nextPeriodOfTDMdata
		for num in tables:
			periodTableCONC[num] = period
			
			if int(period[-2:]) == 1:
				year = str(int(period[:-2]) - 1)
				nextMon = '12'
			else:
				year = period[:-2]
				nextMon = str(int(period[-2:]) - 1)
				if len(nextMon) == 1:
					nextMon = '0' + nextMon
			period = year + str(nextMon)

		def set_period(row):
			num = int(row['TABLENAME'].split('_')[-1])
			val = periodTableCONC[num]
			return val
		# print(MASTER_DATEN)
		MASTER_DATEN['PERIOD'] = MASTER_DATEN.apply(set_period, axis=1)
		# print(MASTER_DATEN)
		MASTER_DATEN.insert(1, 'PERIOD', MASTER_DATEN.pop('PERIOD'))
		# print(MASTER_DATEN)

		# Load MASTER to DB
		self.logger.info(f'\t\t\tDROPPING TABLE SRC_Austria_Customs.dbo.DATEN_TEMP')
		cursor.execute(f"IF OBJECT_ID('SRC_Austria_Customs.dbo.DATEN_TEMP') IS NOT NULL DROP TABLE [SRC_Austria_Customs].[dbo].[DATEN_TEMP];")
		cursor.commit()

		# CREATE TABLE
		create_statements_cols = ''
		for col in MASTER_DATEN.columns:
			if MASTER_DATEN.columns[-1] == col:
				create_statements_cols = create_statements_cols + f'[{col}] [varchar](max) NULL'
			else: 
				create_statements_cols = create_statements_cols + f'[{col}] [varchar](max) NULL,'
		
		self.logger.info(f'\t\t\tCREATING TABLE SRC_Austria_Customs.dbo.DATEN_TEMP')
		cursor.execute(f"""
		CREATE TABLE SRC_Austria_Customs.dbo.DATEN_TEMP(
			{create_statements_cols}
		)
		""")
		cursor.commit()

		# Loop through list of split df and insert into table
		self.logger.info(f'\t\t\tINSERTING {len(MASTER_DATEN)} ROWS INTO SRC_Austria_Customs.dbo.DATEN_TEMP')
		self.logger.info('')
		insert_to_temp_table = f'INSERT INTO SRC_Austria_Customs.dbo.DATEN_TEMP VALUES(?,?,?,?,?,?,?,?,?)'
		cursor.fast_executemany = True

		# INSERT INTO TEMP
		masterValues = MASTER_DATEN.values.tolist()
		rowStepper = 1000000
		# Inserting 1000000 rows of data at a time to prevent error
		for rowI in range(0,len(masterValues), rowStepper):
			cursor.executemany(insert_to_temp_table, masterValues[rowI:rowI+rowStepper])
			cursor.commit()

		cursorMaria.close()
		connMaria.close()
		cursor.close()
		conn.close()
		self.logger.info('\tData Loaded.')

	def processData(self):
		self.logger.info('\tProcessing Data...')
		# Grab SQL processing script
		sqlScriptPath = "Y:\\_PyScripts\\Damon\\Austria_Customs\\Austria_Customs_pyscript.sql"
		sql = self.getSql(sqlScriptPath)
		# Create Connection
		conn = pyodbc.connect(
			driver = '{ODBC Driver 17 for SQL Server}', 
			server = 'SF-PROCESS',
			database = 'SRC_Austria_Customs', 
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
				FROM [SRC_Austria_Customs].[DBO].[RUNNING-STATUS];
				''').fetchone()
			
			except:
				continue
			STATUS = query[0]
			if STATUS == 0:
				self.logger.info('\t\t\tSQL Script Executed.')
				running = False 
				cursor.execute('''
				IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[SRC_Austria_Customs].[dbo].[RUNNING-STATUS]') AND type in (N'U'))
				DROP TABLE [SRC_Austria_Customs].[DBO].[RUNNING-STATUS];
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
			database = 'SRC_Macao', 
			uid = 'sa', 
			pwd = 'Harpua88'
		)
		cursor = conn.cursor()

		# Generic Checks
		self.logger.info('\t\tGeneric Checks...')

		PRE_FINAL_TABLE_EXP = '[SRC_Austria_Customs].[dbo].[SRC_STEP10_EXP]'
		PRE_FINAL_TABLE_IMP = '[SRC_Austria_Customs].[dbo].[SRC_STEP10_IMP]'

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

		EXP_TOTALS = pd.read_sql_query(f'''
		SELECT PERIOD, 
			CAST(SUM(VALUE_EURO) AS VARCHAR) AS [VALUE_EURO], 
			CAST(SUM(VALUE) AS VARCHAR) AS [VALUE], 
			CAST(SUM(QTY1) AS VARCHAR) AS [QTY1], 
			CAST(SUM(QTY2) AS VARCHAR) AS [QTY2]
		FROM {PRE_FINAL_TABLE_EXP}
		GROUP BY PERIOD
		ORDER BY 1 DESC;
		''',conn)
		# print(EXP_TOTALS)
		IMP_TOTALS = pd.read_sql_query(f'''
		SELECT PERIOD, 
			CAST(SUM(VALUE_EURO) AS VARCHAR) AS [VALUE_EURO], 
			CAST(SUM(VALUE) AS VARCHAR) AS [VALUE], 
			CAST(SUM(QTY1) AS VARCHAR) AS [QTY1], 
			CAST(SUM(QTY2) AS VARCHAR) AS [QTY2]
		FROM {PRE_FINAL_TABLE_IMP}
		GROUP BY PERIOD
		ORDER BY 1 DESC;
		''',conn)
		# print(IMP_TOTALS)

		# Check if newest period exists within final tables before inserting, want a table of length zero 
		check_E8 = pd.read_sql_query(f"""
		SELECT DISTINCT a.PERIOD, b.PERIOD
		FROM {PRE_FINAL_TABLE_EXP} a
		LEFT JOIN [SRC_Austria_Customs].[dbo].[E8] b
		ON a.PERIOD = b.PERIOD
		WHERE b.PERIOD IS NOT NULL
		AND a.PERIOD = '{self.nextPeriodOfTDMdata}';
		""", conn)
		# print(len(check_E8))
		check_I8 = pd.read_sql_query(f"""
		SELECT DISTINCT a.PERIOD, b.PERIOD
		FROM {PRE_FINAL_TABLE_IMP} a
		LEFT JOIN [SRC_Austria_Customs].[dbo].[I8] b
		ON a.PERIOD = b.PERIOD
		WHERE b.PERIOD IS NOT NULL
		AND a.PERIOD = '{self.nextPeriodOfTDMdata}';
		""", conn)
		# print(len(check_I8))

		insrtCondition = False
		if list(genericChecks.all())[0] and len(check_E8) == 0 and len(check_I8) == 0:
			# INSERT INTO ARCHIVE
			cursor.execute("""
		DELETE FROM [SRC_Austria_Customs].[dbo].[Archive_SRC_EXP]
		WHERE [PERIOD] IN
			(
			SELECT DISTINCT [PERIOD]
			FROM [SRC_Austria_Customs].[dbo].[SRC_EXP]
			);

		INSERT INTO [SRC_Austria_Customs].[dbo].[Archive_SRC_EXP]
		SELECT *
		FROM [SRC_Austria_Customs].[dbo].[SRC_EXP];

		DELETE FROM [SRC_Austria_Customs].[dbo].[Archive_SRC_IMP]
		WHERE [PERIOD] IN
			(
			SELECT DISTINCT [PERIOD]
			FROM [SRC_Austria_Customs].[dbo].[SRC_IMP]
			);

		INSERT INTO [SRC_Austria_Customs].[dbo].[Archive_SRC_IMP]
		SELECT *
		FROM [SRC_Austria_Customs].[dbo].[SRC_IMP];
		
			""") 
			cursor.commit()

			# INSERT INTO FINAL TABLES
			cursor.execute('''
			-- INSERT EXPORTS
			DELETE FROM [SRC_Austria_Customs].[dbo].[E8]
			WHERE [PERIOD] IN
			(
			SELECT DISTINCT [PERIOD]
			FROM [SRC_Austria_Customs].[dbo].[SRC_STEP10_EXP]
			);
			INSERT INTO [SRC_Austria_Customs].[dbo].[E8]
			SELECT [CTY_RPT]
				,[CTY_PTN]
				,[PERIOD]
				,[YR]
				,[MON]
				,[COMMODITY]
				,[VALUE_EURO]
				,[VALUE]
				,[UNIT1]
				,[QTY1]
				,[UNIT2]
				,[QTY2]
			FROM [SRC_Austria_Customs].[dbo].[SRC_STEP10_EXP];
		
			-- INSERT IMPORTS
			DELETE FROM [SRC_Austria_Customs].[dbo].[I8]
			WHERE [PERIOD] IN
			(
			SELECT DISTINCT [PERIOD]
			FROM [SRC_Austria_Customs].[dbo].[SRC_STEP10_IMP]
			);
			INSERT INTO [SRC_Austria_Customs].[dbo].[I8]
			SELECT [CTY_RPT]
				,[CTY_PTN]
				,[PERIOD]
				,[YR]
				,[MON]
				,[COMMODITY]
				,[VALUE_EURO]
				,[VALUE]
				,[UNIT1]
				,[QTY1]
				,[UNIT2]
				,[QTY2]
			FROM [SRC_Austria_Customs].[dbo].[SRC_STEP10_IMP];
			
			''')
			cursor.commit()

			# drop tables
			cursor.execute('''
			--DROP TABLE [SRC_Austria_Customs].[dbo].[SRC_EXP_STEP1];
			--DROP TABLE [SRC_Austria_Customs].[dbo].[SRC_IMP_STEP1];
			DROP TABLE [SRC_Austria_Customs].[dbo].[SRC_STEP2];
			DROP TABLE [SRC_Austria_Customs].[dbo].[SRC_STEP3];
			DROP TABLE [SRC_Austria_Customs].[dbo].[SRC_STEP4];
			DROP TABLE [SRC_Austria_Customs].[dbo].[SRC_STEP5];
			DROP TABLE [SRC_Austria_Customs].[dbo].[SRC_STEP6];
			DROP TABLE [SRC_Austria_Customs].[dbo].[SRC_STEP7];
			DROP TABLE [SRC_Austria_Customs].[dbo].[SRC_STEP8];
			DROP TABLE [SRC_Austria_Customs].[dbo].[SRC_STEP9];
			DROP TABLE [SRC_Austria_Customs].[dbo].[SRC_STEP10_EXP];
			DROP TABLE [SRC_Austria_Customs].[dbo].[SRC_STEP10_IMP];
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

	at = ETL('Austria_Customs')
	at.run()

