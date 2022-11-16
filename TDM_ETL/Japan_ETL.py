import logging
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
import pandas_profiling
import io
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
d = datetime.today().weekday()


class Japan_ETL():

	def __init__(self):
		conn = pyodbc.connect(
							'Driver={SQL Server};'
							'Server=SF-TEST\SFTESTDB;'
							'Database=Control;'
							'UID=sa;'
							'PWD=Harpua88;'
							'Trusted_Connection=No;'
		)

		Q = """
		SELECT StopYM
		from [Control].[dbo].[Data_Availability_Monthly]
		WHERE Declarant = 'Japan'"""
		# int(list(pd.read_sql_query(Q, conn)['StopYM'])[0][-2:])
		self.lastPeriodOfTDMdata = list(pd.read_sql_query(Q, conn)['StopYM'])[0]
		# self.lastPeriodOfTDMdata = '202208' # hard set for testing
		
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
		
		self.current_year = datetime.now().year
		self.logPath = "Y:\\_PyScripts\\Damon\\Japan\\Log"
		self.downloadPath = "Y:\\_PyScripts\\Damon\\Japan\\Downloads"
		self.archivePath = "Y:\\Japan\\All Data Files"
		self.excelBugPath = 'C:\\Users\\DDiaz.ANH\\AppData\\Local\\Temp\\gen_py\\3.10\\00020813-0000-0000-C000-000000000046x0x1x9'
		if os.path.exists(self.excelBugPath):
			shutil.rmtree(self.excelBugPath)
		if not os.path.exists(self.logPath):
			os.makedirs(self.logPath)
		if not os.path.exists(self.downloadPath):
			os.makedirs(self.downloadPath)
		if not os.path.exists(self.archivePath):
			os.makedirs(self.archivePath)

		for file in os.listdir(self.downloadPath): 
			if os.path.isfile(os.path.join(self.downloadPath, file)):
				os.remove(os.path.join(self.downloadPath, file))
		
		# print(self.lastPeriodOfTDMdata)
		self.YEARLY_TOTAL_PATH = f'Y:\\Japan\\Check totals\\{self.lastPeriodOfTDMdata}\\YEARLY_TOTAL.csv'
		# print(self.YEARLY_TOTAL_PATH)
		self.MONTHLY_TOTAL_PATH = f'Y:\\Japan\\Check totals\\{self.lastPeriodOfTDMdata}\\MONTHLY_TOTAL.csv'
		# print(self.MONTHLY_TOTAL_PATH)
		self.newDataPath = "Y:\\Japan\\Data Files\\Newest Month"
		
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
		self.logger.info('Japan_ETL')
		self.logger.info(f'Last Period Published: {self.lastPeriodOfTDMdata}')

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
		message["Subject"] = f'Japan {phase}'
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

		chrome_options = webdriver.ChromeOptions()
		prefs = {'download.default_directory':self.downloadPath}
		chrome_options.add_experimental_option('prefs', prefs)
		self.driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=chrome_options)
		self.driver.get(URL)
	
	def extractTotals(self):
		self.logger.info('Extracting Totals...')
		try:
			self.connect('https://www.customs.go.jp/toukei/srch/indexe.htm?M=27&P=0')
			self.driver.switch_to_frame("FR_M_INFO")
			self.driver.switch_to_frame("FR_DISP")
			# search YEARLY
			WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="contents"]/div/p[1]/input[1]')))
			self.driver.execute_script('parent.FR_CTRL.js27s_onSubmit();')
			# grab YEARLY table
			self.driver.switch_to_frame("FR_M_INFO")
			YEARLY_TOTAL = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="resultwidth"]'))).get_attribute('innerHTML')
			YEARLY_URL = self.driver.current_url
			# go back to search form
			self.driver.execute_script('FcReSearch();')
			# select MONTHLY
			self.driver.switch_to.default_content()
			self.driver.switch_to_frame("FR_M_INFO")
			self.driver.switch_to_frame("FR_DISP")
			WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="contents"]/div/p[3]/select[1]/option[3]'))).click()
			# search MONTHLY
			WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="contents"]/div/p[1]/input[1]')))
			self.driver.execute_script('parent.FR_CTRL.js27s_onSubmit();')
			# # grab MONTHLY table
			self.driver.switch_to_frame("FR_M_INFO")
			MONTHLY_TOTAL = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="resultwidth"]'))).get_attribute('innerHTML')
			MONTHLY_URL = self.driver.current_url
			# convert html to df
			YEARLY_TOTAL = pd.read_html(YEARLY_TOTAL)[1]
			# print(YEARLY_TOTAL)
			MONTHLY_TOTAL = pd.read_html(MONTHLY_TOTAL)[1]
			MONTHLY_TOTAL[MONTHLY_TOTAL.columns[0]] = MONTHLY_TOTAL[MONTHLY_TOTAL.columns[0]].str.replace('/', '')
			# print(MONTHLY_TOTAL)

			YEARLY_TOTAL.to_csv(self.YEARLY_TOTAL_PATH)
			MONTHLY_TOTAL.to_csv(self.MONTHLY_TOTAL_PATH)
		except Exception as e:
			self.logger.exception(e)
			if self.driver:
				if self.driver.session_id is not None:
					self.driver.quit()
			self.extractTotals()
		self.driver.quit()
		

	def extractData(self):
		self.logger.info('\tExtracting Data...')
		dataToLoad = False
		for file in os.listdir(self.downloadPath): 
			if os.path.isfile(os.path.join(self.downloadPath, file)):
				os.remove(os.path.join(self.downloadPath, file))
		
		self.linksToScrape = {
			'Export':f'https://www.e-stat.go.jp/en/stat-search/files?page=1&query={self.yearToQuery}&layout=dataset&toukei=00350300&tstat=000001013141&cycle=1&tclass1=000001013180&tclass2=000001013181&tclass3val=0&metadata=1&data=1',
			'Import':f'https://www.e-stat.go.jp/en/stat-search/files?page=1&query={self.yearToQuery}&layout=dataset&toukei=00350300&tstat=000001013141&cycle=1&tclass1=000001013180&tclass2=000001013182&tclass3val=0&metadata=1&data=1',
			'Export_ADF':f'https://www.e-stat.go.jp/en/stat-search/files?page=1&query={self.yearToQuery}&layout=dataset&toukei=00350300&tstat=000001013144&cycle=1&tclass1=000001013244&tclass2=000001013245&tclass3val=0&metadata=1&data=1',
			'Import_ADF':f'https://www.e-stat.go.jp/en/stat-search/files?page=1&query={self.yearToQuery}&layout=dataset&toukei=00350300&tstat=000001013144&cycle=1&tclass1=000001013244&tclass2=000001013246&tclass3val=0&metadata=1&data=1'
		}
		downloadFolderContent = len([x for x in os.listdir(self.downloadPath) if zipfile.is_zipfile(self.downloadPath + f'\\{x}')])
		# print(downloadFolderContent)
		for flow, URL in self.linksToScrape.items():
			# print(flow)
			# print(URL)
			self.connect(URL)
			try:
				updatedDate = WebDriverWait(self.driver,60).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/main/div[2]/section/div[2]/div/div/div[1]/section/section/div/div[4]/div/article[1]/div/ul/li[4]'))).text				
			except:
				self.logger.info('Updated Date Not Found...')
				raise ValueError('Error: Date not found. Not able to make comparison.')
			updatedDate = updatedDate.split(' ')[-1]
			# updatedDate = '2022-10-31' # hard set for testing
			if datetime.now().strftime('%Y-%m-%d') != updatedDate:
				self.logger.info('Updated Date does Not Equal Current Date')
				self.logger.info(f'Current Date: {datetime.now().strftime("%Y-%m-%d")} | Updated Date:{updatedDate}')
				self.driver.quit() # quit browser before exiting script
				sys.exit(0)
			else:
				self.logger.info(f'Possibly New Data...{datetime.now().strftime("%Y-%m-%d")} = {updatedDate} ')
				dataToLoad = True
			# ###########################################################################
			# click batch download
			# WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/main/div[2]/section/div[2]/div/div/div[1]/section/section/div/div[2]/div'))).click()
			# confirm download
			# WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/main/div[2]/section/div[2]/div/div/div[4]/div/div[2]/div[4]/button[2]'))).click()
			
			# https://www.e-stat.go.jp/en/stat-search/files/data?page=1&files=000009162593:1,000009162615:1,000009162635:1,000009162639:1,000009162641:1,000009162589:1,000009162591:1,000009162595:1,000009162597:1,000009162599:1,000009162601:1,000009162603:1,000009162605:1,000009162607:1,000009162609:1,000009162611:1,000009162613:1,000009162631:1,000009162633:1,000009162637:1,000009162617:1,000009162619:1,000009162621:1,000009162623:1,000009162625:1,000009162627:1,000009162629:1
			# click batch download
			sleep(8)
			self.driver.execute_script("$('.js-dl_all').click()")
			sleep(3)
			# confirm download
			self.driver.execute_script("$('.js-all-file-download').click()")
			isDownloaded = False
			while isDownloaded == False:
				sleep(3)
				if downloadFolderContent < len([x for x in os.listdir(self.downloadPath) if zipfile.is_zipfile(self.downloadPath + f'\\{x}')]):
					downloadFolderContent = len([x for x in os.listdir(self.downloadPath) if zipfile.is_zipfile(self.downloadPath + f'\\{x}')])
					isDownloaded = True
			self.driver.quit()
			list_of_files = glob.glob(f'{self.downloadPath}//*') # * means all if need specific format then *.csv
			latest_file = max(list_of_files, key=os.path.getctime)
			os.rename(f'{latest_file}', self.downloadPath + f'\\{flow}.zip')
	
		
		for folder in os.listdir(self.downloadPath):
			shutil.copy2(self.downloadPath + f'\\{folder}', self.newDataPath)
			path = Path(self.newDataPath + f'\\{folder}')
			path.rename(f'{self.newDataPath}\\{self.nextPeriodOfTDMdata}{folder}')

		self.StatusEmail('New Data', f'Possibly New Japan Data', '')
		self.logger.info('\tData Extraction Complete.')
		return dataToLoad
		
	def readData(self):
		self.logger.info('\tReading Data...')
		self.MASTER = pd.DataFrame()
		for zipFolder in os.listdir(self.newDataPath):
			zipFolderPath = self.newDataPath + f'\\{zipFolder}'
			# print(self.newDataPath)
			# print(zipFolderPath)
			
			zf = zipfile.ZipFile(zipFolderPath)
			# print(zf.namelist())

			for csvFileI in range(1, len(zf.namelist())):
				csvFile = zf.namelist()[csvFileI]
				if len(csvFile) > 0:
					# print(csvFile)
					with zf.open(zf.namelist()[csvFileI]) as csvObj:
						df = pd.read_csv(csvObj)
						# print(df)
						self.MASTER = pd.concat([self.MASTER, df], ignore_index=True)
			try:
				if not os.path.exists(f'{self.archivePath}\\{self.nextPeriodOfTDMdata}'):
					os.makedirs(f'{self.archivePath}\\{self.nextPeriodOfTDMdata}')
				
				shutil.copy2(zipFolderPath, f'{self.archivePath}\\{self.nextPeriodOfTDMdata}\\'+zipFolder)
			except Exception as e:
				self.logger.info('Error moving files to archive.')
				self.logger.info(e)
			zf.close()
			os.remove(os.path.join(self.newDataPath, zipFolder))
		
		tempDF = self.MASTER.drop('Exp or Imp', axis=1)
		tempDF = tempDF.drop('Year', axis=1)
		tempDF = tempDF.drop('HS', axis=1)
		tempDF = tempDF.drop('Country', axis=1)
		tempDF = tempDF.drop('Unit1', axis=1)
		tempDF = tempDF.drop('Unit2', axis=1)
		tempDF = tempDF.drop('Quantity1-Year', axis=1)
		tempDF = tempDF.drop('Quantity2-Year', axis=1)
		tempDF = tempDF.drop('Value-Year', axis=1)
		tempDF = tempDF.drop('Custom', axis=1)
		# print(tempDF)
		
		cols_with_all_zeros = []
		# grab month  
		for col in reversed(tempDF.columns):
		
			# print(col)
			tempCol = list(tempDF[col])
			result = tempCol.count(tempCol[0]) == len(tempCol)

			if result:
				# all values are equal
				cols_with_all_zeros.append(col.split('-')[1])
			else:
				# not all values are equal
				pass
		# print(cols_with_all_zeros)
		
		months = set([i for i in cols_with_all_zeros if cols_with_all_zeros.count(i)==3])
		# print(months)
		num_months = []
		for month in months:
			month_obj = datetime.strptime(month, '%b')
			month_num = month_obj.month
			num_months.append(month_num)
		nextMonthOfData = min(num_months)
		
		lastMonth = int(self.lastPeriodOfTDMdata[-2:])
		# print(self.lastPeriodOfTDMdata)
		# print(nextMonthOfData)
		self.logger.info(f'\t\t\tlast Month Of TDM data:{lastMonth}')
		self.logger.info(f'\t\t\tJapan Site Month Of Data:{nextMonthOfData-1}')
		if lastMonth == nextMonthOfData - 1:
			isDataNew = False
		else:
			isDataNew = True
			self.StatusEmail('New Data', 'New Data Downloaded', '', False)

		self.logger.info('\tData Read.')
		return isDataNew

	def loadData(self):
		self.logger.info('\tLoading Data...')
		self.MASTER = pd.DataFrame()
		for zipFolder in os.listdir(self.downloadPath):
			zipFolderPath = self.downloadPath + f'\\{zipFolder}'
			# print(self.downloadPath)
			# print(zipFolderPath)
			
			zf = zipfile.ZipFile(zipFolderPath)
			# print(zf.namelist())

			for csvFileI in range(1, len(zf.namelist())):
				csvFile = zf.namelist()[csvFileI]
				if len(csvFile) > 0:
					# print(csvFile)
					with zf.open(zf.namelist()[csvFileI]) as csvObj:
						df = pd.read_csv(csvObj)
						df.insert(0, 'FILENAME', f'{zipFolderPath}\\{csvFile}')
						# print(df)
						self.MASTER = pd.concat([self.MASTER, df], ignore_index=True)
		# seperates ADF data from REG data by finding where 'Custom' field is empty
		# print(self.MASTER)
		
		dfs = {}
		self.REG = self.MASTER[self.MASTER['Custom'].isna()]
		self.REG = self.REG.drop([self.REG.columns[-1]], axis=1)
		# self.REG.insert(len(self.REG.columns), 'load_date', datetime.now())
		dfs['REG'] = self.REG
		# print(self.REG)
		self.ADF = self.MASTER[self.MASTER['Custom'].notna()]
		customCOL = self.ADF.pop('Custom')
		self.ADF.insert(2, 'Custom', customCOL)
		# self.ADF.insert(len(self.ADF.columns), 'load_date', datetime.now())
		dfs['ADF'] = self.ADF
		# print(self.ADF)
		
		# create connection to [SRC_Japan].[dbo].[TEMP]
		conn = pyodbc.connect(
			driver = '{ODBC Driver 17 for SQL Server}', 
			server = 'SEVENFARMS_DB1', 
			database = 'SRC_Japan', 
			uid = 'sa', 
			pwd = 'Harpua88')

		cursor = conn.cursor()
		for k,v in dfs.items():
			v = v.astype(str)
			self.logger.info(f'\t\t\tDROPPING TABLE SRC_Japan.dbo.TEMP_{k}')
			cursor.execute(f"IF OBJECT_ID('SRC_Japan.dbo.TEMP_{k}') IS NOT NULL DROP TABLE [SRC_Japan].[dbo].[TEMP_{k}];")
			cursor.commit()	

			# CREATE TABLE 
			create_statements_cols = ''
			for col in v.columns: 
				if v.columns[-1] == col:
					create_statements_cols = create_statements_cols + f'[{col}] [varchar](max) NULL'
				else:
					create_statements_cols = create_statements_cols + f'[{col}] [varchar](max) NULL,'
			
			self.logger.info(f'\t\t\tCREATING TABLE SRC_Japan.dbo.TEMP_{k}')
			cursor.execute(f"""
			CREATE TABLE SRC_Japan.dbo.TEMP_{k}(
				{create_statements_cols}
			)
			""")
			cursor.commit()
			
			# Loop through list of split df and insert into table
			self.logger.info(f'\t\t\tINSERTING INTO SRC_Japan.dbo.TEMP_{k}')
			self.logger.info('')
			insert_to_REG_table = f'INSERT INTO TEMP_{k} VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'
			insert_to_ADF_table = f'INSERT INTO TEMP_{k} VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'
			cursor.fast_executemany = True
			# for split_df in np.array_split(v, 10):
			# print(split_df)
			# INSERT INTO TEMP
			if k == 'REG':
				cursor.executemany(insert_to_REG_table, v.values.tolist())
			else:
				cursor.executemany(insert_to_ADF_table, v.values.tolist())
			cursor.commit()
	
		cursor.close()
		conn.close()
		self.logger.info('\tData Loaded.')
	
	def processData(self):
		self.logger.info('Processing Data...')
		sqlScriptPath = "Y:\\_PyScripts\\Damon\\Japan\\JAPAN_pyscript.sql"
		# Grabs SQL Script
		sql = self.getSql(sqlScriptPath)
		# Create Connection
		conn = pyodbc.connect(
			driver = '{ODBC Driver 17 for SQL Server}', 
			server = 'SEVENFARMS_DB1', 
			database = 'SRC_Japan', 
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
				FROM [SRC_Japan].[DBO].[RUNNING-STATUS];
				''').fetchone()

			except: 
				continue
			STATUS = query[0]
			if STATUS == 0: 
				self.logger.info('\t\t\tSQL Script Executed.')
				running = False
				cursor.execute('''
				IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[SRC_Japan].[dbo].[RUNNING-STATUS]') AND type in (N'U'))
				DROP TABLE [SRC_Japan].[DBO].[RUNNING-STATUS];
				''')
				cursor.commit()
			else: 
				running = True
		cursor.close()
		conn.close()
		self.logger.info('\tData Processed.')

	def dataChecks(self):
		self.logger.info('\tPerforming Checks...')
		# Grab TOTALS 
		YEARLY_TOTALS = pd.read_csv(self.YEARLY_TOTAL_PATH)
		MONTHLY_TOTALS = pd.read_csv(self.MONTHLY_TOTAL_PATH)
		offenders = []

		# Create Connection
		conn = pyodbc.connect(
			driver = '{ODBC Driver 17 for SQL Server}', 
			server = 'SEVENFARMS_DB1', 
			database = 'SRC_Japan', 
			uid = 'sa', 
			pwd = 'Harpua88')
		cursor = conn.cursor()
		
		# Generic checks
		self.logger.info('\tGeneric Checks...')
		dfE8 = pd.read_sql_query('''
		SELECT * 
		FROM [SRC_Japan].[dbo].[TEMP_REG_EXP_STEP7]
		''', conn)
		dfI8 = pd.read_sql_query('''
		SELECT * 
		FROM [SRC_Japan].[dbo].[TEMP_REG_IMP_STEP7]
		''', conn)
		failedChecks = {'CHECK PERFORMED': ['STATUS'],
						'E8 NULL VALUES DO NOT EXIST':[False], 
						'I8 NULL VALUES DO NOT EXIST':[False], 
						'E8 COMMODITY_LEN IS 9 DIGITS':[False],
						'I8 COMMODITY_LEN IS 9 DIGITS':[False],
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

		if dfE8['COMMODITY'].astype(str).map(len).nunique() == 1 and dfE8['COMMODITY'].astype(str).map(len).unique()[0] == 9:
			failedChecks['E8 COMMODITY_LEN IS 9 DIGITS'][0] = True

		if dfI8['COMMODITY'].astype(str).map(len).nunique() == 1 and dfI8['COMMODITY'].astype(str).map(len).unique()[0] == 9:
			failedChecks['I8 COMMODITY_LEN IS 9 DIGITS'][0] = True

		# Check CTY_PTN 
		tempE8 = pd.read_sql_query('''
		SELECT DISTINCT CTY_PTN, CTY_ABBR
		FROM [SRC_Japan].[dbo].[TEMP_REG_EXP_STEP7] A
		LEFT JOIN [SP_MASTER].[dbo].[CTY_MASTER] B
		ON A.CTY_PTN = B.CTY_ABBR
		''', conn)
		if tempE8.isnull().values.any() != True: 
			failedChecks['ALL E8 CTY_PTN EXIST IN CTY_MASTER DB1'][0] = True

		tempI8 = pd.read_sql_query('''
		SELECT DISTINCT CTY_PTN, CTY_ABBR
		FROM [SRC_Japan].[dbo].[TEMP_REG_IMP_STEP7] A
		LEFT JOIN [SP_MASTER].[dbo].[CTY_MASTER] B
		ON A.CTY_PTN = B.CTY_ABBR
		''', conn)
		if tempI8.isnull().values.any() != True: 
			failedChecks['ALL I8 CTY_PTN EXIST IN CTY_MASTER DB1'][0] = True

		# Check UNIT1
		tempE8 = pd.read_sql_query('''
		SELECT DISTINCT UNIT1, ANH_UOM
		FROM [SRC_Japan].[dbo].[TEMP_REG_EXP_STEP7] A
		LEFT JOIN [SP_MASTER].[dbo].[UOM_MASTER] B
		ON A.UNIT1 = B.ANH_UOM
		''', conn)
		if tempE8.isnull().values.any() != True: 
			failedChecks['ALL E8 UNIT1 EXIST IN UOM_MASTER DB1'][0] = True

		tempI8 = pd.read_sql_query('''
		SELECT DISTINCT UNIT1, ANH_UOM
		FROM [SRC_Japan].[dbo].[TEMP_REG_IMP_STEP7] A
		LEFT JOIN [SP_MASTER].[dbo].[UOM_MASTER] B
		ON A.UNIT1 = B.ANH_UOM
		''', conn)
		if tempI8.isnull().values.any() != True: 
			failedChecks['ALL I8 UNIT1 EXIST IN UOM_MASTER DB1'][0] = True

		# Check UNIT2 
		tempE8 = pd.read_sql_query('''
		SELECT DISTINCT UNIT2, ANH_UOM
		FROM [SRC_Japan].[dbo].[TEMP_REG_EXP_STEP7] A
		LEFT JOIN [SP_MASTER].[dbo].[UOM_MASTER] B
		ON A.UNIT2 = B.ANH_UOM
		''', conn)
		if tempE8.isnull().values.any() != True: 
			failedChecks['ALL E8 UNIT2 EXIST IN UOM_MASTER DB1'][0] = True

		tempI8 = pd.read_sql_query('''
		SELECT DISTINCT UNIT2, ANH_UOM
		FROM [SRC_Japan].[dbo].[TEMP_REG_IMP_STEP7] A
		LEFT JOIN [SP_MASTER].[dbo].[UOM_MASTER] B
		ON A.UNIT2 = B.ANH_UOM
		''', conn)
		if tempI8.isnull().values.any() != True: 
			failedChecks['ALL I8 UNIT2 EXIST IN UOM_MASTER DB1'][0] = True

		# print(failedChecks)
		genericChecks = pd.DataFrame.from_dict(failedChecks)
		genericChecks = genericChecks.set_index('CHECK PERFORMED').transpose()
		# print(genericChecks)

		# Generic ADF Checks
		dfADFE8 = pd.read_sql_query('''
		SELECT * 
		FROM [SRC_Japan].[dbo].[TEMP_ADF_EXP_STEP7]
		''', conn)
		dfADFI8 = pd.read_sql_query('''
		SELECT * 
		FROM [SRC_Japan].[dbo].[TEMP_ADF_IMP_STEP7]
		''', conn)
		failedADFChecks = {'CHECK ADF PERFORMED': ['STATUS'],
						'E8 ADF NULL VALUES DO NOT EXIST':[False], 
						'I8 ADF NULL VALUES DO NOT EXIST':[False], 
						'E8 ADF COMMODITY_LEN IS 9 DIGITS':[False],
						'I8 ADF COMMODITY_LEN IS 9 DIGITS':[False],
						'ALL E8 ADF CTY_PTN EXIST IN CTY_MASTER DB1': [False],
						'ALL I8 ADF CTY_PTN EXIST IN CTY_MASTER DB1': [False],
						'ALL E8 ADF UNIT1 EXIST IN UOM_MASTER DB1': [False],
						'ALL I8 ADF UNIT1 EXIST IN UOM_MASTER DB1': [False],
						'ALL E8 ADF UNIT2 EXIST IN UOM_MASTER DB1': [False],
						'ALL I8 ADF UNIT2 EXIST IN UOM_MASTER DB1': [False],
		}

		if dfADFE8.isnull().values.any() != True:
			failedADFChecks['E8 ADF NULL VALUES DO NOT EXIST'][0] = True

		if dfADFI8.isnull().values.any() != True:
			failedADFChecks['I8 ADF NULL VALUES DO NOT EXIST'][0] = True

		if dfADFE8['COMMODITY'].astype(str).map(len).nunique() == 1 and dfADFE8['COMMODITY'].astype(str).map(len).unique()[0] == 9:
			failedADFChecks['E8 ADF COMMODITY_LEN IS 9 DIGITS'][0] = True

		if dfADFI8['COMMODITY'].astype(str).map(len).nunique() == 1 and dfADFI8['COMMODITY'].astype(str).map(len).unique()[0] == 9:
			failedADFChecks['I8 ADF COMMODITY_LEN IS 9 DIGITS'][0] = True

		# Check CTY_PTN 
		tempADFE8 = pd.read_sql_query('''
		SELECT DISTINCT CTY_PTN, CTY_ABBR
		FROM [SRC_Japan].[dbo].[TEMP_ADF_EXP_STEP7] A
		LEFT JOIN [SP_MASTER].[dbo].[CTY_MASTER] B
		ON A.CTY_PTN = B.CTY_ABBR
		''', conn)
		if tempADFE8.isnull().values.any() != True: 
			failedADFChecks['ALL E8 ADF CTY_PTN EXIST IN CTY_MASTER DB1'][0] = True

		tempADFI8 = pd.read_sql_query('''
		SELECT DISTINCT CTY_PTN, CTY_ABBR
		FROM [SRC_Japan].[dbo].[TEMP_ADF_IMP_STEP7] A
		LEFT JOIN [SP_MASTER].[dbo].[CTY_MASTER] B
		ON A.CTY_PTN = B.CTY_ABBR
		''', conn)
		if tempADFI8.isnull().values.any() != True: 
			failedADFChecks['ALL I8 ADF CTY_PTN EXIST IN CTY_MASTER DB1'][0] = True

		# Check UNIT1
		tempADFE8 = pd.read_sql_query('''
		SELECT DISTINCT UNIT1, ANH_UOM
		FROM [SRC_Japan].[dbo].[TEMP_ADF_EXP_STEP7] A
		LEFT JOIN [SP_MASTER].[dbo].[UOM_MASTER] B
		ON A.UNIT1 = B.ANH_UOM
		''', conn)
		if tempADFE8.isnull().values.any() != True: 
			failedADFChecks['ALL E8 ADF UNIT1 EXIST IN UOM_MASTER DB1'][0] = True

		tempADFI8 = pd.read_sql_query('''
		SELECT DISTINCT UNIT1, ANH_UOM
		FROM [SRC_Japan].[dbo].[TEMP_ADF_IMP_STEP7] A
		LEFT JOIN [SP_MASTER].[dbo].[UOM_MASTER] B
		ON A.UNIT1 = B.ANH_UOM
		''', conn)
		if tempADFI8.isnull().values.any() != True: 
			failedADFChecks['ALL I8 ADF UNIT1 EXIST IN UOM_MASTER DB1'][0] = True

		# Check UNIT2 
		tempADFE8 = pd.read_sql_query('''
		SELECT DISTINCT UNIT2, ANH_UOM
		FROM [SRC_Japan].[dbo].[TEMP_ADF_EXP_STEP7] A
		LEFT JOIN [SP_MASTER].[dbo].[UOM_MASTER] B
		ON A.UNIT2 = B.ANH_UOM
		''', conn)
		if tempADFE8.isnull().values.any() != True: 
			failedADFChecks['ALL E8 ADF UNIT2 EXIST IN UOM_MASTER DB1'][0] = True

		tempADFI8 = pd.read_sql_query('''
		SELECT DISTINCT UNIT2, ANH_UOM
		FROM [SRC_Japan].[dbo].[TEMP_ADF_IMP_STEP7] A
		LEFT JOIN [SP_MASTER].[dbo].[UOM_MASTER] B
		ON A.UNIT2 = B.ANH_UOM
		''', conn)
		if tempADFI8.isnull().values.any() != True: 
			failedADFChecks['ALL I8 ADF UNIT2 EXIST IN UOM_MASTER DB1'][0] = True

		# print(failedADFChecks)
		genericADFChecks = pd.DataFrame.from_dict(failedADFChecks)
		genericADFChecks = genericADFChecks.set_index('CHECK ADF PERFORMED').transpose()
		# print(genericADFChecks)
		
		# TOTAL CHECKS
		pd.set_option('display.float_format', lambda x: '%.2f' % x)
		# Read All REG TOTAL Tables 
		TEMP_REG_EXP_STEP4_YEAR = pd.read_sql_query('SELECT * FROM TEMP_REG_EXP_STEP4_YEAR ORDER BY 2', conn)
		TEMP_REG_EXP_STEP4_MON = pd.read_sql_query('SELECT * FROM TEMP_REG_EXP_STEP4_MON ORDER BY 2', conn)
		TEMP_REG_IMP_STEP4_YEAR = pd.read_sql_query('SELECT * FROM TEMP_REG_IMP_STEP4_YEAR ORDER BY 2', conn)
		TEMP_REG_IMP_STEP4_MON = pd.read_sql_query('SELECT * FROM TEMP_REG_IMP_STEP4_MON ORDER BY 2', conn)
		# Read All REG STEP7 Tables 
		TEMP_REG_EXP_STEP7 = pd.read_sql_query('SELECT * FROM TEMP_REG_EXP_STEP7', conn)
		TEMP_REG_IMP_STEP7 = pd.read_sql_query('SELECT * FROM TEMP_REG_IMP_STEP7', conn)

		# Read All ADF Tables 
		TEMP_ADF_EXP_STEP4_YEAR = pd.read_sql_query('SELECT * FROM TEMP_ADF_EXP_STEP4_YEAR', conn)
		TEMP_ADF_EXP_STEP4_MON = pd.read_sql_query('SELECT * FROM TEMP_ADF_EXP_STEP4_MON', conn)
		TEMP_ADF_IMP_STEP4_YEAR = pd.read_sql_query('SELECT * FROM TEMP_ADF_IMP_STEP4_YEAR', conn)
		TEMP_ADF_IMP_STEP4_MON = pd.read_sql_query('SELECT * FROM TEMP_ADF_IMP_STEP4_MON', conn)
		# Read All ADF STEP7 Tables 
		TEMP_ADF_EXP_STEP7 = pd.read_sql_query('SELECT * FROM TEMP_ADF_EXP_STEP7', conn)
		TEMP_ADF_IMP_STEP7 = pd.read_sql_query('SELECT * FROM TEMP_ADF_IMP_STEP7', conn)

		# Compare totals Before Processing and After Processing
		dfTotalsCheck_YEARLY = sqldf('''
		SELECT 
		A.YR
		,A.TOTAL_VALUE AS [EXP_TOTAL_PROCESSED]
		,CAST(B.EXPORT AS DECIMAL) AS [EXP_TOTAL_RAW]
		,C.TOTAL_VALUE AS [IMP_TOTAL_PROCESSED]
		,CAST(B.IMPORT AS DECIMAL) AS [IMP_TOTAL_RAW]
		,CAST(B.EXPORT AS DECIMAL) - CAST(A.TOTAL_VALUE AS DECIMAL) AS [EXP_DIFF]
		,CAST(B.IMPORT AS DECIMAL) - CAST(C.TOTAL_VALUE AS DECIMAL) AS [IMP_DIFF]
		,(CAST(B.EXPORT AS DECIMAL) - CAST(A.TOTAL_VALUE AS DECIMAL)) /CAST(B.EXPORT AS DECIMAL) *100 AS [EXP_DIFF PERCENT]
		,(CAST(B.IMPORT AS DECIMAL) - CAST(C.TOTAL_VALUE AS DECIMAL)) /CAST(B.IMPORT AS DECIMAL) *100 AS [IMP_DIFF PERCENT]
		FROM TEMP_REG_EXP_STEP4_YEAR A
		LEFT JOIN TEMP_REG_IMP_STEP4_YEAR C
		ON C.YR = A.YR
		LEFT JOIN YEARLY_TOTALS B
		ON A.YR = B.YEAR
		ORDER BY 1
		''')
		# print(dfTotalsCheck_YEARLY)
		dfTotalsCheck_MONTHLY = sqldf('''
		SELECT 
		CAST(A.YR as text) || CAST(A.MON as text) AS [PERIOD]
		,A.TOTAL_VALUE AS [EXP_TOTAL_PROCESSED]
		,CAST(B.EXPORT AS DECIMAL) AS [EXP_TOTAL_RAW]
		,C.TOTAL_VALUE AS [IMP_TOTAL_PROCESSED]
		,CAST(B.IMPORT AS DECIMAL) AS [IMP_TOTAL_RAW]
		,CAST(B.EXPORT AS DECIMAL) - CAST(A.TOTAL_VALUE AS DECIMAL) AS [EXP_DIFF]
		,CAST(B.IMPORT AS DECIMAL) - CAST(C.TOTAL_VALUE AS DECIMAL) AS [IMP_DIFF]
		,(CAST(B.EXPORT AS DECIMAL) - CAST(A.TOTAL_VALUE AS DECIMAL)) /CAST(B.EXPORT AS DECIMAL) *100 AS [EXP_DIFF PERCENT]
		,(CAST(B.IMPORT AS DECIMAL) - CAST(C.TOTAL_VALUE AS DECIMAL)) /CAST(B.IMPORT AS DECIMAL) *100 AS [IMP_DIFF PERCENT]
		FROM TEMP_REG_EXP_STEP4_MON A
		LEFT JOIN TEMP_REG_IMP_STEP4_MON C
		ON C.MON = A.MON
		LEFT JOIN MONTHLY_TOTALS B
		ON CAST(A.YR as text) || CAST(A.MON as text) = B.[YEAR / MONTH]
		ORDER BY 1
		''')
		# print(dfTotalsCheck_MONTHLY)
		dfTotalsCheck_ADFYEARLY = sqldf('''
		SELECT 
		A.YR
		,A.TOTAL_VALUE AS [EXP_TOTAL_PROCESSED]
		,CAST(B.EXPORT AS DECIMAL) AS [EXP_TOTAL_RAW]
		,C.TOTAL_VALUE AS [IMP_TOTAL_PROCESSED]
		,CAST(B.IMPORT AS DECIMAL) AS [IMP_TOTAL_RAW]
		,CAST(B.EXPORT AS DECIMAL) - CAST(A.TOTAL_VALUE AS DECIMAL) AS [EXP_DIFF]
		,CAST(B.IMPORT AS DECIMAL) - CAST(C.TOTAL_VALUE AS DECIMAL) AS [IMP_DIFF]
		,(CAST(B.EXPORT AS DECIMAL) - CAST(A.TOTAL_VALUE AS DECIMAL)) /CAST(B.EXPORT AS DECIMAL) *100 AS [EXP_DIFF PERCENT]
		,(CAST(B.IMPORT AS DECIMAL) - CAST(C.TOTAL_VALUE AS DECIMAL)) /CAST(B.IMPORT AS DECIMAL) *100 AS [IMP_DIFF PERCENT]
		FROM TEMP_ADF_EXP_STEP4_YEAR A
		LEFT JOIN TEMP_ADF_IMP_STEP4_YEAR C
		ON C.YR = A.YR
		LEFT JOIN YEARLY_TOTALS B
		ON A.YR = B.YEAR
		ORDER BY 1
		''')
		# print(dfTotalsCheck_ADFYEARLY)
		dfTotalsCheck_ADFMONTHLY = sqldf('''
		SELECT 
		CAST(A.YR as text) || CAST(A.MON as text) AS [PERIOD]
		,A.TOTAL_VALUE AS [EXP_TOTAL_PROCESSED]
		,CAST(B.EXPORT AS DECIMAL) AS [EXP_TOTAL_RAW]
		,C.TOTAL_VALUE AS [IMP_TOTAL_PROCESSED]
		,CAST(B.IMPORT AS DECIMAL) AS [IMP_TOTAL_RAW]
		,CAST(B.EXPORT AS DECIMAL) - CAST(A.TOTAL_VALUE AS DECIMAL) AS [EXP_DIFF]
		,CAST(B.IMPORT AS DECIMAL) - CAST(C.TOTAL_VALUE AS DECIMAL) AS [IMP_DIFF]
		,(CAST(B.EXPORT AS DECIMAL) - CAST(A.TOTAL_VALUE AS DECIMAL)) /CAST(B.EXPORT AS DECIMAL) *100 AS [EXP_DIFF PERCENT]
		,(CAST(B.IMPORT AS DECIMAL) - CAST(C.TOTAL_VALUE AS DECIMAL)) /CAST(B.IMPORT AS DECIMAL) *100 AS [IMP_DIFF PERCENT]
		FROM TEMP_ADF_EXP_STEP4_MON A
		LEFT JOIN TEMP_ADF_IMP_STEP4_MON C
		ON C.MON = A.MON
		LEFT JOIN MONTHLY_TOTALS B
		ON CAST(A.YR as text) || CAST(A.MON as text) = B.[YEAR / MONTH]
		ORDER BY 1
		''')
		# print(dfTotalsCheck_ADFMONTHLY)

		# Insert Final Tables Into [SRC_Japan].[dbo].[Exports_NEW] and [SRC_Japan].[dbo].[Imports_NEW]
		# print(list(genericChecks.all())[0])
		if list(genericChecks.all())[0] and list(genericADFChecks.all())[0]:
			cursor.execute('''
			-- EXPORTS 
				DELETE FROM [SRC_Japan].[dbo].[Exports_NEW]
				WHERE PERIOD IN (
				SELECT DISTINCT [PERIOD]
				FROM [SRC_JAPAN].[dbo].[TEMP_REG_EXP_STEP7]
				);
			   INSERT INTO [SRC_Japan].[dbo].[Exports_NEW]
			 SELECT  cast([CTY_RPT] AS VARCHAR(3)) AS [CTY_RPT]
				,cast([COMMODITY] AS VARCHAR(9)) AS [COMMODITY]
			 	,cast([CTY_PTN] AS VARCHAR(7)) AS [CTY_PTN]
			 	,cast([YR] AS Int) AS [YR]
			 	,cast([MON] AS TinyInt) AS [MON]
				,cast([PERIOD] AS VARCHAR(6)) AS [PERIOD]
			 	,cast([VALUE] AS DECIMAL(38,8)) AS [VALUE]
			 	,cast([VALUE_JPY] AS DECIMAL(38,8)) AS [VALUE_JPY]
			 	,cast([UNIT1] AS VARCHAR(3)) AS [UNIT1]
			 	,cast([QTY1] AS DECIMAL(38,2)) AS [QTY1]
			 	,cast([UNIT2] AS VARCHAR(3)) AS [UNIT2]
			 	,cast([QTY2] AS DECIMAL(38,2)) AS [QTY2]
			 FROM [SRC_JAPAN].[dbo].[TEMP_REG_EXP_STEP7];
			-- IMPORTS
				DELETE FROM [SRC_Japan].[dbo].[Imports_NEW]
				WHERE PERIOD IN (
				SELECT DISTINCT [PERIOD]
				FROM [SRC_JAPAN].[dbo].[TEMP_REG_IMP_STEP7]
				);
			 INSERT INTO [SRC_Japan].[dbo].[Imports_NEW]
			 SELECT  cast([CTY_RPT] AS VARCHAR(3)) AS [CTY_RPT]
			 	,cast([CTY_PTN] AS VARCHAR(7)) AS [CTY_PTN]
			 	,cast([COMMODITY] AS VARCHAR(9)) AS [COMMODITY]
			 	,cast([PERIOD] AS VARCHAR(6)) AS [PERIOD]
			 	,cast([YR] AS Int) AS [YR]
			 	,cast([MON] AS TinyInt) AS [MON]
				,cast([VALUE] AS DECIMAL(38,8)) AS [VALUE]
			 	,cast([VALUE_JPY] AS DECIMAL(38,8)) AS [VALUE_JPY]
			 	,cast([UNIT1] AS VARCHAR(3)) AS [UNIT1]
			 	,cast([QTY1] AS DECIMAL(18,2)) AS [QTY1]
			 	,cast([UNIT2] AS VARCHAR(3)) AS [UNIT2]
			 	,cast([QTY2] AS DECIMAL(18,2)) AS [QTY2]
			 FROM [SRC_JAPAN].[dbo].[TEMP_REG_IMP_STEP7];
			''')
			cursor.commit()

			cursor.execute('''
			-- EXPORTS 
			DELETE FROM [SRC_Japan].[dbo].[E8_ADF]
			WHERE PERIOD IN (
			SELECT DISTINCT [PERIOD]
			FROM [SRC_JAPAN].[dbo].[TEMP_ADF_EXP_STEP7]
			);
			INSERT INTO [SRC_Japan].[dbo].[E8_ADF]
			SELECT  cast([CTY_RPT] AS VARCHAR(3)) AS [CTY_RPT]
				,cast([CTY_PTN] AS VARCHAR(7)) AS [CTY_PTN]
				,cast([COMMODITY] AS VARCHAR(9)) AS [COMMODITY]
				,cast([PERIOD] AS VARCHAR(6)) AS [PERIOD]
				,cast([YR] AS Int) AS [YR]
				,cast([MON] AS TinyInt) AS [MON]
				,cast([VALUE] AS DECIMAL(38,8)) AS [VALUE]
				,cast([VALUE_JPY] AS DECIMAL(38,8)) AS [VALUE_JPY]
				,cast([UNIT1] AS VARCHAR(3)) AS [UNIT1]
				,cast([QTY1] AS DECIMAL(16,0)) AS [QTY1]
				,cast([UNIT2] AS VARCHAR(3)) AS [UNIT2]
				,cast([QTY2] AS DECIMAL(16,0)) AS [QTY2]
				,CUSTOM
				,CUSTOMBRANCH
			FROM [SRC_JAPAN].[dbo].[TEMP_ADF_EXP_STEP7];

			-- IMPORTS
			DELETE FROM [SRC_Japan].[dbo].[I8_ADF]
			WHERE PERIOD IN (
			SELECT DISTINCT [PERIOD]
			FROM [SRC_JAPAN].[dbo].[TEMP_ADF_IMP_STEP7]
			);
			INSERT INTO [SRC_Japan].[dbo].[I8_ADF]
			SELECT  cast([CTY_RPT] AS VARCHAR(3)) AS [CTY_RPT]
				,cast([CTY_PTN] AS VARCHAR(7)) AS [CTY_PTN]
				,cast([COMMODITY] AS VARCHAR(9)) AS [COMMODITY]
				,cast([PERIOD] AS VARCHAR(6)) AS [PERIOD]
				,cast([YR] AS Int) AS [YR]
				,cast([MON] AS TinyInt) AS [MON]
				,cast([VALUE] AS DECIMAL(38,8)) AS [VALUE]
				,cast([VALUE_JPY] AS DECIMAL(38,8)) AS [VALUE_JPY]
				,cast([UNIT1] AS VARCHAR(3)) AS [UNIT1]
				,cast([QTY1] AS DECIMAL(16,0)) AS [QTY1]
				,cast([UNIT2] AS VARCHAR(3)) AS [UNIT2]
				,cast([QTY2] AS DECIMAL(16,0)) AS [QTY2]
				,CUSTOM
				,CUSTOMBRANCH
			FROM [SRC_JAPAN].[dbo].[TEMP_ADF_IMP_STEP7];
			''')
			cursor.commit()

			# DROP REG TABLES GRABBED
			cursor.execute('''
			DROP TABLE [SRC_JAPAN].[dbo].[TEMP_REG_EXP];
			DROP TABLE [SRC_JAPAN].[dbo].[TEMP_REG_EXP_STEP4_YEAR];
			DROP TABLE [SRC_JAPAN].[dbo].[TEMP_REG_EXP_STEP4_MON];
			DROP TABLE [SRC_JAPAN].[dbo].[TEMP_REG_IMP];
			DROP TABLE [SRC_JAPAN].[dbo].[TEMP_REG_IMP_STEP4_YEAR];
			DROP TABLE [SRC_JAPAN].[dbo].[TEMP_REG_IMP_STEP4_MON];
			--DROP TABLE [SRC_JAPAN].[dbo].[TEMP_REG_EXP_STEP7];
			--DROP TABLE [SRC_JAPAN].[dbo].[TEMP_REG_IMP_STEP7];
			''')
			cursor.commit()
			# DROP ADF TABLES GRABBED
			cursor.execute('''
			DROP TABLE [SRC_JAPAN].[dbo].[TEMP_ADF_EXP];
			DROP TABLE [SRC_JAPAN].[dbo].[TEMP_ADF_EXP_STEP4_YEAR];
			DROP TABLE [SRC_JAPAN].[dbo].[TEMP_ADF_EXP_STEP4_MON];
			DROP TABLE [SRC_JAPAN].[dbo].[TEMP_ADF_IMP];
			DROP TABLE [SRC_JAPAN].[dbo].[TEMP_ADF_IMP_STEP4_YEAR];
			DROP TABLE [SRC_JAPAN].[dbo].[TEMP_ADF_IMP_STEP4_MON];
			--DROP TABLE [SRC_JAPAN].[dbo].[TEMP_ADF_EXP_STEP7];
			--DROP TABLE [SRC_JAPAN].[dbo].[TEMP_ADF_IMP_STEP7];
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
			insrtCondition = False
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
		message["Subject"] = f'Japan Check'
		message["from"] = "TDM Data Team"
		message["To"] = ", ".join(recipients)

		html = f"""\
		<html>
			<body>
				<table style="width:100%">
					<tbody>
						<tr>
							<td>{genericChecks.to_html().replace('<td>True', '<td><p class = "grn">True').replace('<td>False', '<td><p class = "rd">False').replace('</td>', '</p></td>')}</td>
							<td>{genericADFChecks.to_html().replace('<td>True', '<td><p class = "grn">True').replace('<td>False', '<td><p class = "rd">False').replace('</td>', '</p></td>')}</td> 
						</tr>
					</tbody>
				</table>
				<br>
			
				<hr>
				<h3>Totals Check: YEARLY</h3>
					{dfTotalsCheck_YEARLY.to_html(index=False)}
				<br>
				<hr>
				<h3>ADF Totals Check: YEARLY</h3>
					{dfTotalsCheck_ADFYEARLY.to_html(index=False)}
				<br>
				<hr>
				<h3>Totals Check: MONTHLY</h3>
					{dfTotalsCheck_MONTHLY.to_html(index=False)}
				<br>
				<hr>
				<h3>ADF Totals Check: MONTHLY</h3>
					{dfTotalsCheck_ADFMONTHLY.to_html(index=False)}
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
		self.logger.info('\tChecks Performed.')
	
	def autoUnitCheck(self, conn, cursor): 
		self.logger.info('\t\tWorking on Unit Checks...')
		if not os.path.exists(rf'Y:\Japan\Unit Checks\{self.nextPeriodOfTDMdata}'):
			os.mkdir(rf'Y:\Japan\Unit Checks\{self.nextPeriodOfTDMdata}')
		f_path = Path(rf'Y:\Japan\Unit Checks\{self.nextPeriodOfTDMdata}')  # file located somewhere else
		table_matches = {
			'Exports_NEW':'TEMP_REG_EXP_STEP7',
			'Imports_NEW':'TEMP_REG_EXP_STEP7',
			'E8_ADF':'TEMP_ADF_EXP_STEP7', 
			'I8_ADF':'TEMP_ADF_IMP_STEP7', 
		}

		for k,v in table_matches.items():
			
			if os.path.exists(self.excelBugPath):
				shutil.rmtree(self.excelBugPath)
			file_name = f'{k[0]}8_{k[-3:]}.xlsx'
			self.logger.info(f'\t\t\t{file_name}')

			# [SRC_Japan].[dbo].[Imports_NEW]
			pivotData = pd.read_sql_query(f'''
			SELECT PERIOD, COUNT(DISTINCT commodity) AS NbrCommods,UNIT2,sum(qty2) as QTY2,sum(value) as USD,CTY_RPT AS CTY_RPT
			, SUM(QTY1) AS QTY1, UNIT1,SUBSTRING(COMMODITY,1,2) AS CH
			FROM [SRC_Japan].[dbo].[{k}]
			GROUP BY PERIOD,unit1,SUBSTRING(COMMODITY,1,2) ,unit2,cty_rpt
			ORDER BY period
			''', conn)
			
			pivotData.to_excel(f_path/file_name, sheet_name=f'{k[0]}8', index=False)
			self.generateUnitChecks(f_path, file_name, f'{k[0]}8')
			if os.path.exists(self.excelBugPath):
				shutil.rmtree(self.excelBugPath)
		self.logger.info('\t\t\tUnit Checks Generated.')
		
		return conn, cursor

	def run(self):
		self.logger.info('Launching Japan ETL...')

		# Data Extraction
		try:
			dataToLoad = self.extractData()
		except Exception as e:
			self.logger.exception('extractData() Error')
			self.StatusEmail('extractData() Error', e, traceback.format_exc())
			sys.exit()
		# dataToLoad = True # hard set for testing
		if dataToLoad:
			# Compare Data to confirm it is new
			try:
				newData = self.readData()
			except Exception as e:
				self.logger.exception('readData() Error')
				self.StatusEmail('readData() Error', e, traceback.format_exc())
				sys.exit()
		
			# newData = True
			if newData:
				self.logger.info('\t\t*New Data Available*')

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

				# Extract Totals
				try:
					self.extractTotals()
				except Exception as e:
					self.logger.exception('extractTotals() Error')
					self.StatusEmail('extractTotals() Error', e, traceback.format_exc())
					sys.exit()

				# Data Checks
				try:
					self.dataChecks()
				except Exception as e: 
					self.logger.exception('dataChecks() Error')
					self.StatusEmail('dataChecks() Error', e, traceback.format_exc())
					sys.exit()

		else: 
			self.logger.info('\t\tNo New Data is Available.')


		self.logger.info('Japan ETL Finished.')


if __name__ == '__main__':
	
	ja = Japan_ETL()
	ja.run()