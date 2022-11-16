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
import urllib.parse
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
from selenium.webdriver.support.select import Select
win32c = win32.constants


class ETL():
	def __init__(self, country):
		self.country = country 
		
		conn = pyodbc.connect(
							'Driver={SQL Server};'
							'Server=SEVENFARMS_DB3;'
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
		#self.nextMon = '06' # hard set for testing
		
		self.current_year = datetime.now().year
		self.today = datetime.now().strftime('%#m/%#d/%Y')
		# URLS
		self.data_source_url = "http://www.dsec.gov.mo/EMTS/NCEMSearchData.aspx?Type=2"
		self.data_source_url_en = "http://www.dsec.gov.mo/EMTS/NCEMSearchData.aspx?lang=en-US&Type=2"
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
		self.logger.info('Macao_ETL')
	
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

	def connect(self, URL):
		self.logger.info(f'\t\tConnecting {URL}')

		chrome_options = webdriver.ChromeOptions()
		prefs = {'download.default_directory':self.downloadPath}
		chrome_options.add_experimental_option('prefs', prefs)
		# chrome_options.add_argument('--headless')
		self.driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=chrome_options)
		self.driver.get(URL)
	
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
		#query ctys 16 to 255

		# create session
		s = requests.Session()
		# navigate to form
		r_form = s.get(self.data_source_url)
		soup = bs(r_form.text)
		sitePeriod = soup.find_all('span', {'id':f'plcRoot_Layout_zoneLeft_SearchByCodeOrCountry{self.current_year}_LiteralRevision'})[0].text
		# print(sitePeriod)
		
		# check if data is new 
		parsedSitePeriod = "".join([char for char in sitePeriod if char.isdigit()])[:-4]
		if len(parsedSitePeriod) == 5:
			parsedSitePeriod = parsedSitePeriod[:4] + '0' + parsedSitePeriod[4:]
			self.logger.info(f'\t\tLatest Date of Source Data: {parsedSitePeriod}')
		else:
			self.logger.info(f'\t\tLatest Date of Source Data: {parsedSitePeriod}')
		
		if parsedSitePeriod != self.nextPeriodOfTDMdata:
			self.logger.info(f'\t\t\tLooking for {self.nextPeriodOfTDMdata}')
			self.logger.info('\t\t\tNo New Data Available.')
			s.close()
			dataToLoad = False
			return dataToLoad
		else:
			dataToLoad = True
			self.logger.info('\t\t\t*New Data Available.')
			self.StatusEmail('New Data', f'New {self.country} Data', '')
			s.close()

			# loop through countries
			for ctyI in range(0, 255+1):
			# for ctyI in range(224, 224+1):
				if ctyI < 16 and ctyI > 0:
					continue
				self.logger.info(f'\t\t\tQuery Country Index:{ctyI}')
				# create new english session
				s = requests.Session()
				# navigate to form
				r_form = s.get(self.data_source_url_en)
				soup = bs(r_form.text)
				# print(soup)
				# input()
				# os.system('cls')

				# post NCMESearchData
				__VIEWSTATE = soup.find_all('input', {'id':'__VIEWSTATE'})[0]['value']
				# print(__VIEWSTATE)
				payload={
				'__LASTFOCUS':'',
				'__EVENTTARGET':f'plcRoot$Layout$zoneLeft$SearchByCodeOrCountry{self.current_year}$RadMenu1',
				'__EVENTARGUMENT':f'0:{ctyI}',
				'__VIEWSTATE':__VIEWSTATE,
				'lng':'en-US',
				'__VIEWSTATEGENERATOR':'A5343185',
				f'plcRoot$Layout$zoneLeft$SearchByCodeOrCountry{self.current_year}$rdChoice':0,
				f'plcRoot_Layout_zoneLeft_SearchByCodeOrCountry{self.current_year}_RadMenu1_ClientState':'',
				f'plcRoot$Layout$zoneLeft$SearchByCodeOrCountry{self.current_year}$ddlYear':self.yearToQuery,
				f'plcRoot$Layout$zoneLeft$SearchByCodeOrCountry{self.current_year}$ddlStartMonth':int(self.nextMon),
				f'plcRoot$Layout$zoneLeft$SearchByCodeOrCountry{self.current_year}$ddlStopMonth':int(self.nextMon)
				}
				headers = {
				'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
				'Accept-Language': 'en-US,en;q=0.9',
				'Cache-Control': 'max-age=0',
				'Connection': 'keep-alive',
				'Content-Type': 'application/x-www-form-urlencoded',
				'Cookie': f'CMSPreferredCulture=en-US; ASP.NET_SessionId={r_form.cookies["ASP.NET_SessionId"]}; VisitorStatus=1; ViewMode=0; _ga=GA1.3.2028629558.1660666131; _gid=GA1.3.730426954.1660666131; CMSPreferredCulture=zh-MO; CurrentVisitStatus=2; ViewMode=0; VisitorStatus=1',
				'Origin': 'http://www.dsec.gov.mo',
				'Referer': 'http://www.dsec.gov.mo/EMTS/NCEMSearchData.aspx?lang=en-US&Type=2',
				'Upgrade-Insecure-Requests': '1',
				'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36'
				}
				# request country
				self.logger.info('\t\t\t\tPOST: Country Selection...')
				r_country_sel = s.request("POST", self.data_source_url_en, headers=headers, data=payload)
				soup = bs(r_country_sel.text)
				# print(soup)
				ctyName = soup.select('.rmLink')[0]
				if os.path.exists(self.downloadPath + f'\\{ctyName.text}_{self.nextPeriodOfTDMdata}_Totals.csv') or os.path.exists(self.downloadPath + f'\\{ctyName.text}_{self.nextPeriodOfTDMdata}_Skipped.csv'):
					continue
				
				# parse
				__VIEWSTATE = soup.find_all('input', {'id':'__VIEWSTATE'})[0]['value']

				# post NCMESearchData
				payload={
				'__LASTFOCUS':'',
				'__EVENTTARGET':'',
				'__EVENTARGUMENT':'',
				'__VIEWSTATE':__VIEWSTATE,
				'lng':'en-US',
				'__VIEWSTATEGENERATOR':'A5343185',
				f'plcRoot$Layout$zoneLeft$SearchByCodeOrCountry{self.current_year}$rdChoice':0,
				f'plcRoot_Layout_zoneLeft_SearchByCodeOrCountry{self.current_year}_RadMenu1_ClientState':'',
				f'plcRoot$Layout$zoneLeft$SearchByCodeOrCountry{self.current_year}$ddlYear':self.yearToQuery,
				f'plcRoot$Layout$zoneLeft$SearchByCodeOrCountry{self.current_year}$ddlStartMonth':int(self.nextMon),
				f'plcRoot$Layout$zoneLeft$SearchByCodeOrCountry{self.current_year}$ddlStopMonth':int(self.nextMon),
				f'plcRoot$Layout$zoneLeft$SearchByCodeOrCountry{self.current_year}$btnSearchByCountry':'Search'
				}
				# request country totals
				self.logger.info('\t\t\t\tPOST: Country Totals...')
				r_totals = s.request("POST", self.data_source_url_en, headers=headers, data=payload)
				
				soup = bs(r_totals.text)
				# print(soup)
				# input()
				# parse 
				__VIEWSTATE = soup.find_all('input', {'id':'__VIEWSTATE'})[0]['value']
				cty = soup.find_all('span', {'id':f'plcRoot_Layout_zoneLeft_SearchByCountry{self.current_year}_lblSearchCountry'})[0].text
				if ctyI == 0:
					totals_table = soup.find_all('table', {'id':f'plcRoot_Layout_zoneLeft_SearchByCountry{self.current_year}_GridView2'})[0]
				else:
					totals_table = soup.find_all('table', {'id':f'plcRoot_Layout_zoneLeft_SearchByCountry{self.current_year}_GridView1'})[0]
				TOTALS_DF = pd.read_html(str(totals_table))[0]
				# delete first row
				TOTALS_DF = TOTALS_DF.iloc[2: , 2:]
				TOTALS_DF = TOTALS_DF.rename(columns={
					TOTALS_DF.columns[0]:'Trade type',
					TOTALS_DF.columns[1]:'Macao Pataca',
					TOTALS_DF.columns[2]:'Weight (KG)',
				})
				# print(TOTALS_DF)
				if ctyI == 0:
					# print(TOTALS_DF)
					TOTALS_DF.drop(columns=TOTALS_DF.columns[3], axis=1, inplace=True)
					TOTALS_DF.drop(columns=TOTALS_DF.columns[3], axis=1, inplace=True)
					TOTALS_DF.drop(columns=TOTALS_DF.columns[4], axis=1, inplace=True)
					TOTALS_DF.drop(columns=TOTALS_DF.columns[4], axis=1, inplace=True)
					TOTALS_DF.drop(columns=TOTALS_DF.columns[2], axis=1, inplace=True)
					TOTALS_DF = TOTALS_DF.rename(columns={
					TOTALS_DF.columns[0]:'Trade type',
					TOTALS_DF.columns[1]:'Macao Pataca',
					TOTALS_DF.columns[2]:'Weight (KG)',
					})
					# print(TOTALS_DF)
				
				# check if df has data
				if len(TOTALS_DF) == 0:
					self.logger.info(f'\t\t\t\t\tNo {cty} Data')
					self.logger.info(f'\t\t\t\t\t\tSkipping {cty}...')
					skippedDF = pd.DataFrame()
					skippedDF.to_csv(self.downloadPath + f'\\{cty}_{self.nextPeriodOfTDMdata}_Skipped.csv', index=False)
					# post SearchByCountry New Search
					payload = {
						'__EVENTTARGET':'',
						'__EVENTARGUMENT':'',
						'__LASTFOCUS':'',
						'__VIEWSTATE':__VIEWSTATE,
						'lng':'en-US',
						'__VIEWSTATEGENERATOR':'A5343185',
						'__VIEWSTATEENCRYPTED':'',
						'plcRoot$Layout$zoneLeft$SearchByCountry2022$btnNewSearch':'New Search',
					}
					headers = {
					'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
					'Accept-Language': 'en-US,en;q=0.9',
					'Cache-Control': 'max-age=0',
					'Connection': 'keep-alive',
					'Content-Type': 'application/x-www-form-urlencoded',
					'Cookie': f'CMSPreferredCulture=en-US; ASP.NET_SessionId={r_form.cookies["ASP.NET_SessionId"]}; ViewMode=0; _ga=GA1.3.2028629558.1660666131; _gid=GA1.3.730426954.1660666131; VisitorStatus=2; CurrentVisitStatus=2; _gat=1; CMSPreferredCulture=zh-MO; VisitorStatus=1',
					'Origin': 'http://www.dsec.gov.mo',
					'Referer': 'http://www.dsec.gov.mo/EMTS/SearchByCountry.aspx',
					'Upgrade-Insecure-Requests': '1',
					'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36'
					}
					self.logger.info('\t\t\t\tPOST: Returning to Main Menu...')
					r_newSearch = s.request('POST', "http://www.dsec.gov.mo/EMTS/SearchByCountry.aspx", headers=headers, data=payload)
					soup = bs(r_newSearch.text)
					# print(soup)
					# input()
					continue
				else:
					self.logger.info(f'\t\t\t\t\tExtracting {cty} Totals...')
					
					# determine flows
					flows = totals_table.find_all('input', {'value':'rbCode'})
					# print(inputs)
					if ctyI != 0:
						# loop through flows
						for f in flows:
							
							# print(f['id'])
							# print(f['name])

							# click on flow
							# post SearchByCountry
							payload={
							'__EVENTTARGET':f['id'],
							'__EVENTARGUMENT':'',
							'__LASTFOCUS':'',
							'__VIEWSTATE':__VIEWSTATE,
							'lng':'en-US',
							'__VIEWSTATEGENERATOR':'A5343185',
							'__VIEWSTATEENCRYPTED':'',
							f['name']:'rbCode',
							}
							headers = {
							'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
							'Accept-Language': 'en-US,en;q=0.9',
							'Cache-Control': 'max-age=0',
							'Connection': 'keep-alive',
							'Content-Type': 'application/x-www-form-urlencoded',
							'Cookie': f'CMSPreferredCulture=en-US; ASP.NET_SessionId={r_form.cookies["ASP.NET_SessionId"]}; ViewMode=0; _ga=GA1.3.2028629558.1660666131; _gid=GA1.3.730426954.1660666131; CurrentVisitStatus=2; VisitorStatus=2; CMSPreferredCulture=zh-MO; VisitorStatus=1',
							'Origin': 'http://www.dsec.gov.mo',
							'Referer': 'http://www.dsec.gov.mo/EMTS/SearchByCountry.aspx',
							'Upgrade-Insecure-Requests': '1',
							'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36'
							}
							# request hs codes
							self.logger.info('\t\t\t\tPOST: Requesting HS Codes...')
							r_rbCode = s.request("POST", "http://www.dsec.gov.mo/EMTS/SearchByCountry.aspx", headers=headers, data=payload)
							soup = bs(r_rbCode.text)
							hs_8_digit = soup.find_all('input', {'id':f"plcRoot_Layout_zoneLeft_SearchByCountryDetail{self.current_year}_btnShowAll", 'value':'HS 8-digit level'})
							hs_4_digit = soup.find_all('input', {'id':f"plcRoot_Layout_zoneLeft_SearchByCountryDetail{self.current_year}_btnShowAll", 'value':'HS 4-digit level'})
							if hs_8_digit:
								digitLevel = 'HS 8-digit level'
							elif hs_4_digit:
								digitLevel = 'HS 4-digit level'
							else:
								self.logger.info(f'\t\t\t\t{cty}: Problem finding HS detailed digit level button')
							# parse 
							__VIEWSTATE = soup.find_all('input', {'id':'__VIEWSTATE'})[0]['value']
							

							# click show all 8 digit commodities
							# post SearchbyCountryDetail
							payload={
							'__EVENTTARGET':'',
							'__EVENTARGUMENT':'',
							'__LASTFOCUS':'',
							'__VIEWSTATE':__VIEWSTATE,
							'lng':'en-US',
							'__VIEWSTATEGENERATOR':'A5343185',
							'__VIEWSTATEENCRYPTED':'',
							f'plcRoot$Layout$zoneLeft$SearchByCountryDetail{self.current_year}$btnShowAll':digitLevel,
							}
							headers = {
							'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
							'Accept-Language': 'en-US,en;q=0.9',
							'Cache-Control': 'max-age=0',
							'Connection': 'keep-alive',
							'Content-Type': 'application/x-www-form-urlencoded',
							'Cookie': f'CMSPreferredCulture=en-US; ASP.NET_SessionId={r_form.cookies["ASP.NET_SessionId"]}; ViewMode=0; _ga=GA1.3.2028629558.1660666131; _gid=GA1.3.730426954.1660666131; CurrentVisitStatus=2; VisitorStatus=2; CMSPreferredCulture=zh-MO; VisitorStatus=1',
							'Origin': 'http://www.dsec.gov.mo',
							'Referer': 'http://www.dsec.gov.mo/EMTS/SearchbyCountryDetail.aspx',
							'Upgrade-Insecure-Requests': '1',
							'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36'
							}
							self.logger.info('\t\t\t\tPOST: Specifying 8 Digit Codes...')
							r_commodities = s.request('POST', 'http://www.dsec.gov.mo/EMTS/SearchbyCountryDetail.aspx', headers=headers, data=payload)
							soup = bs(r_commodities.text)
							# print(soup)
							searchCondition_table = soup.find_all('table')[0]
							searchConditionDF = pd.read_html(str(searchCondition_table))[0]
							# print(searchConditionDF)
							searchVars = list(searchConditionDF[searchConditionDF.columns[2]])
							PTN_CTY = searchVars[1]
							TIME = searchVars[2]
							FLOW = searchVars[-1]
							
							# MASTER dataframe for flow
							MASTER_FLOW = pd.DataFrame()
							pageNum = 1
							# loop through pages
							notLastPage = True
							while notLastPage:
								self.logger.info(f'\t\t\t\t\t\t{PTN_CTY} {FLOW} {TIME} Page:{pageNum}')
								# parse 
								__VIEWSTATE = soup.find_all('input', {'id':'__VIEWSTATE'})[0]['value']
								commodity_table = soup.find_all('table', {'id':f'plcRoot_Layout_zoneLeft_SearchByCountryDetail{self.current_year}_GridView1'})[0]
								# print(commodity_table)
								
								
								
								pageNum += 1
								if len(str(pageNum)) == 2 and int(str(pageNum)[-1]) == 1:
									lastClickablePage = commodity_table.find_all('table')[0].find_all('a')[-1].text
									if lastClickablePage == '...':
										nextPageExists = True
									else: 
										nextPageExists = False
									tempDF = pd.read_html(str(commodity_table))[0].iloc[:-1, :]
									
								else:
									pageTable = commodity_table.find_all('table')
									if pageTable:
										tempDF = pd.read_html(str(commodity_table))[0].iloc[:-1, :]
										nextPageExists = pageTable[0].find_all('td', string=f'{pageNum}')
										
									else:
										tempDF = pd.read_html(str(commodity_table))[0].iloc[:, :]
										nextPageExists = []
										
								# print(tempDF)
								prevDF = tempDF
								MASTER_FLOW = pd.concat([MASTER_FLOW, tempDF], ignore_index=True)
								
								if nextPageExists:
									# post next page of data
									payload={
									'__EVENTTARGET':f'plcRoot$Layout$zoneLeft$SearchByCountryDetail{self.current_year}$GridView1',
									'__EVENTARGUMENT':f'Page${pageNum}',
									'__LASTFOCUS':'',
									'__VIEWSTATE':__VIEWSTATE,
									'lng':'en-US',
									'__VIEWSTATEGENERATOR':'A5343185',
									'__VIEWSTATEENCRYPTED':'',
									}
									headers = {
									'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
									'Accept-Language': 'en-US,en;q=0.9',
									'Cache-Control': 'max-age=0',
									'Connection': 'keep-alive',
									'Content-Type': 'application/x-www-form-urlencoded',
									'Cookie': f'CMSPreferredCulture=en-US; ASP.NET_SessionId={r_form.cookies["ASP.NET_SessionId"]}; ViewMode=0; _ga=GA1.3.2028629558.1660666131; _gid=GA1.3.730426954.1660666131; CurrentVisitStatus=2; VisitorStatus=2; CMSPreferredCulture=zh-MO; VisitorStatus=1',
									'Origin': 'http://www.dsec.gov.mo',
									'Referer': 'http://www.dsec.gov.mo/EMTS/SearchbyCountryDetail.aspx',
									'Upgrade-Insecure-Requests': '1',
									'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36'
									}
									self.logger.info('\t\t\t\tPOST: Paging...')
									r_commodities_nextPage = s.request('POST', 'http://www.dsec.gov.mo/EMTS/SearchbyCountryDetail.aspx', headers=headers, data=payload)
									soup = bs(r_commodities_nextPage.text)
									
								else:
									notLastPage = False
							MASTER_FLOW.to_csv(self.downloadPath + f'\\{PTN_CTY}_{self.nextPeriodOfTDMdata}_{FLOW}.csv', index=False)
							# navigate back to flows
							# equivalent to click back button to 2 digit hs
							headers = {
							'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
							'Accept-Language': 'en-US,en;q=0.9',
							'Connection': 'keep-alive',
							'Cookie': f'CMSPreferredCulture=en-US; ASP.NET_SessionId={r_form.cookies["ASP.NET_SessionId"]}; ViewMode=0; _ga=GA1.3.2028629558.1660666131; _gid=GA1.3.730426954.1660666131; VisitorStatus=2; CurrentVisitStatus=2; _gat=1; CMSPreferredCulture=zh-MO; VisitorStatus=1',
							'Referer': 'http://www.dsec.gov.mo/EMTS/SearchByCountry.aspx',
							'Upgrade-Insecure-Requests': '1',
							'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36'
							}
							self.logger.info('\t\t\t\tGET: Returning to 2 Digit Codes...')
							r = s.request('GET', 'http://www.dsec.gov.mo/EMTS/SearchbyCountryDetail.aspx', headers=headers, data={})
							soup = bs(r.text)
							# print(soup)
							
							# equivalent to click back button to flows
							headers = {
							'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
							'Accept-Language': 'en-US,en;q=0.9',
							'Connection': 'keep-alive',
							'Cookie': f'CMSPreferredCulture=en-US; ASP.NET_SessionId={r_form.cookies["ASP.NET_SessionId"]}; ViewMode=0; _ga=GA1.3.2028629558.1660666131; _gid=GA1.3.730426954.1660666131; VisitorStatus=2; CurrentVisitStatus=2; CMSPreferredCulture=zh-MO; VisitorStatus=1',
							'Referer': 'http://www.dsec.gov.mo/EMTS/NCEMSearchData.aspx?Type=2',
							'Upgrade-Insecure-Requests': '1',
							'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.0.0 Safari/537.36'
							}
							self.logger.info('\t\t\t\tGET: Returning to Flow Menu...')
							r = s.request('GET', "http://www.dsec.gov.mo/EMTS/SearchByCountry.aspx", headers=headers, data={})
							soup = bs(r.text)
							__VIEWSTATE = soup.find_all('input', {'id':'__VIEWSTATE'})[0]['value']
							# print(soup)
					TOTALS_DF.to_csv(self.downloadPath + f'\\{cty}_{self.nextPeriodOfTDMdata}_Totals.csv', index=False)
			allDownloadsComplete = True		
							
		self.logger.info('\tData Extraction Complete.')
		return dataToLoad

	def loadData(self):
		self.logger.info('\tLoading Data...')

		# Create Connection to [SRC_Israel]
		conn = pyodbc.connect(
			driver = '{ODBC Driver 17 for SQL Server}', 
			server = 'SEVENFARMS_DB1',
			database = 'SRC_Macao',
			uid = 'sa',
			pwd = 'Harpua88',
		)
		cursor = conn.cursor()

		# Used to hold all data to be loaded
		MASTER = pd.DataFrame()
		
		for file in os.listdir(self.downloadPath):
			self.logger.info(f'\t\tReading {file}...')
			
			# path to file
			filePath = f'{self.downloadPath}\\{file}'
			# print(filePath.split('_')[-1])
			if filePath.split('_')[-1] != 'Totals.csv' and filePath.split('_')[-1] != 'Skipped.csv':
				tempDF = pd.read_csv(filePath)
				tempDF.insert(0, 'COUNTRY_CODE', '')
				tempDF.insert(0, 'COUNTRY', file.split('_')[0])
				tempDF.insert(0, 'MONTH', f'{self.nextPeriodOfTDMdata[-2:]}')
				tempDF.insert(0, 'YEAR', f'{self.nextPeriodOfTDMdata[:-2]}')
				tempDF.insert(0, 'FLOW', f'{filePath.split(".")[0].split("_")[-1]}')
				tempDF.insert(0, 'FILE_MONTH', f'{self.nextPeriodOfTDMdata[-2:]}')
				tempDF.insert(0, 'FILE_YEAR', f'{self.nextPeriodOfTDMdata[:-2]}')
				tempDF.insert(0, 'FILENAME', f'{self.archivePath}\\{file}')

				# Add tempDF to MASTER
				MASTER = pd.concat([MASTER, tempDF], ignore_index=True)
				
			if file.split('_')[-1] == 'Totals.csv' and file.split('_')[0] != 'ALL COUNTRIES':
				tempDF = pd.read_csv(filePath)
				tempDF.insert(0, 'COUNTRY_CODE', '')
				tempDF.insert(0, 'COUNTRY', file.split('_')[0])
				tempDF.insert(1, 'Commodity description', 'TOTALS')
				tempDF.insert(0, 'MONTH', f'{self.nextPeriodOfTDMdata[-2:]}')
				tempDF.insert(0, 'YEAR', f'{self.nextPeriodOfTDMdata[:-2]}')
				tempDF.insert(0, 'FILE_MONTH', f'{self.nextPeriodOfTDMdata[-2:]}')
				tempDF.insert(0, 'FILE_YEAR', f'{self.nextPeriodOfTDMdata[:-2]}')
				tempDF.insert(0, 'FILENAME', f'{self.archivePath}\\{file}')
				tempDF = tempDF.rename(columns={
					'Trade type':'FLOW'
				})

				# Add tempDF to MASTER
				MASTER = pd.concat([MASTER, tempDF], ignore_index=True)

			if file.split('_')[-1] == 'Totals.csv' and file.split('_')[0] == 'ALL COUNTRIES':
				tempDF = pd.read_csv(filePath)
				tempDF.insert(0, 'COUNTRY_CODE', '')
				tempDF.insert(0, 'COUNTRY', file.split('_')[0])
				tempDF.insert(1, 'Commodity description', 'ALL COUNTRIES')
				tempDF.insert(0, 'MONTH', f'{self.nextPeriodOfTDMdata[-2:]}')
				tempDF.insert(0, 'YEAR', f'{self.nextPeriodOfTDMdata[:-2]}')
				tempDF.insert(0, 'FILE_MONTH', f'{self.nextPeriodOfTDMdata[-2:]}')
				tempDF.insert(0, 'FILE_YEAR', f'{self.nextPeriodOfTDMdata[:-2]}')
				tempDF.insert(0, 'FILENAME', f'{self.archivePath}\\{file}')
				tempDF = tempDF.rename(columns={
					'Trade type':'FLOW'
				})
				
				# Add tempDF to MASTER
				MASTER = pd.concat([MASTER, tempDF], ignore_index=True)

			# moves file to archive once read
			shutil.move(f'{self.downloadPath}\\{file}', f'{self.archivePath}\\{file}')
		self.logger.info('\t\tAll Data Read.')

		MASTER = MASTER.astype(str)
		print(MASTER)

		# Load MASTER to DB
		self.logger.info('\t\t\tDROPPING TABLE SRC_Macao.dbo.TEMP')
		cursor.execute(f"IF OBJECT_ID('SRC_Macao.dbo.TEMP') IS NOT NULL DROP TABLE [SRC_Macao].[dbo].[TEMP];")
		cursor.commit()

		# CREATE TABLE
		create_statements_cols = ''
		for col in MASTER.columns:
			if MASTER.columns[-1] == col:
				create_statements_cols = create_statements_cols + f'[{col}] [varchar](max) NULL'
			else: 
				create_statements_cols = create_statements_cols + f'[{col}] [varchar](max) NULL,'

		self.logger.info(f'\t\t\tCREATING TABLE SRC_Macao.dbo.TEMP')
		cursor.execute(f"""
		CREATE TABLE SRC_Macao.dbo.TEMP(
			{create_statements_cols}
		)
		""")
		cursor.commit()

		# Loop through list of split df and insert into table
		self.logger.info(f'\t\t\tINSERTING {len(MASTER)} ROWS INTO SRC_Macao.dbo.TEMP')
		self.logger.info('')
		insert_to_temp_table = f'INSERT INTO SRC_Macao.dbo.TEMP VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)'
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
		sqlScriptPath = "Y:\\_PyScripts\\Damon\\Macao\\Macao_pyscript.sql"
		sql = self.getSql(sqlScriptPath)
		# Create Connection
		conn = pyodbc.connect(
			driver = '{ODBC Driver 17 for SQL Server}', 
			server = 'SEVENFARMS_DB1',
			database = 'SRC_Macao', 
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
				FROM [SRC_Macao].[DBO].[RUNNING-STATUS];
				''').fetchone()
			
			except:
				continue
			STATUS = query[0]
			if STATUS == 0:
				self.logger.info('\t\t\tSQL Script Executed.')
				running = False 
				cursor.execute('''
				IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[SRC_Macao].[dbo].[RUNNING-STATUS]') AND type in (N'U'))
				DROP TABLE [SRC_Macao].[DBO].[RUNNING-STATUS];
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
			server = 'SEVENFARMS_DB1', 
			database = 'SRC_Macao', 
			uid = 'sa', 
			pwd = 'Harpua88'
		)
		
		cursor = conn.cursor()

		# Generic Checks
		self.logger.info('\t\tGeneric Checks...')

		PRE_FINAL_TABLE_EXP = '[SRC_Macao].[dbo].[TEMP_STEP5_EXP]'
		PRE_FINAL_TABLE_IMP = '[SRC_Macao].[dbo].[TEMP_STEP5_IMP]'

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
		SELECT a.[PERIOD], a.[CTY_PTN]
		, cast(sum(cast([VALUE_MOP] as decimal(38,8))) as varchar) as [PROCESSED VALUE_MOP]
		,[Macao Pataca] AS [RAW VALUE_MOP]
		,sum(cast([VALUE_MOP] as decimal(38,8))) - [Macao Pataca] AS [VALUE DIFF]
		--,cast(sum(cast([QTY1] as decimal(38,8))) as varchar) as [PROCESSED QTY1]
		,cast(sum(cast([QTY2] as decimal(38,8))) as varchar) as [PROCESSED QTY2]
		,[KG] AS [RAW QTY2]
		, sum(cast([QTY2] as decimal(38,8))) - [KG] AS [QTY2 DIFF]
		FROM {PRE_FINAL_TABLE_EXP} a
		LEFT JOIN [SRC_Macao].[dbo].[TEMP_TOTALS_STEP1] b
		on a.[CTY_PTN] = b.[CTY_PTN]
		WHERE b.[FLOW] = 'Exports'
		GROUP BY a.PERIOD, a.[CTY_PTN],[KG], [Macao Pataca]
		ORDER BY 1
		''',conn)
		# print(EXP_TOTALS)
		IMP_TOTALS = pd.read_sql_query(f'''
		SELECT a.[PERIOD], a.[CTY_PTN]
		, cast(sum(cast([VALUE_MOP] as decimal(38,8))) as varchar) as [PROCESSED VALUE_MOP]
		,[Macao Pataca] AS [RAW VALUE_MOP]
		,sum(cast([VALUE_MOP] as decimal(38,8))) - [Macao Pataca] AS [VALUE DIFF]
		--,cast(sum(cast([QTY1] as decimal(38,8))) as varchar) as [PROCESSED QTY1]
		,cast(sum(cast([QTY2] as decimal(38,8))) as varchar) as [PROCESSED QTY2]
		,[KG] AS [RAW QTY2]
		, sum(cast([QTY2] as decimal(38,8))) - [KG] AS [QTY2 DIFF]
		FROM {PRE_FINAL_TABLE_IMP} a
		LEFT JOIN [SRC_Macao].[dbo].[TEMP_TOTALS_STEP1] b
		on a.[CTY_PTN] = b.[CTY_PTN]
		WHERE b.[FLOW] = 'Imports'
		GROUP BY a.PERIOD, a.[CTY_PTN],[KG], [Macao Pataca]
		ORDER BY 1
		''',conn)
		# print(IMP_TOTALS)
		
		def styleCompDF(row):
			highlight = 'background-color: red;'
			default = 'background-color: green;'
			returnVar = [default, default]
			try:
				int(row['VALUE DIFF'])
				int(row['QTY2 DIFF'])

				if int(row['VALUE DIFF']) != 0 and int(row['QTY2 DIFF']) != 0:
					returnVar[0] = highlight
					returnVar[1] = highlight
			except:
				pass
			return returnVar

		IMP_TOTALS = IMP_TOTALS.style.apply(styleCompDF, subset=['VALUE DIFF', 'QTY2 DIFF'], axis=1).hide(axis='index')
		EXP_TOTALS = EXP_TOTALS.style.apply(styleCompDF, subset=['VALUE DIFF', 'QTY2 DIFF'], axis=1).hide(axis='index')

		# Check if newest period exists within final tables before inserting, want a table of length zero 
		check_E8 = pd.read_sql_query(f'''
		SELECT DISTINCT a.PERIOD, b.PERIOD
		FROM {PRE_FINAL_TABLE_EXP} a
		LEFT JOIN [SRC_Macao].[dbo].[E8] b
		ON a.PERIOD = b.PERIOD
		WHERE b.PERIOD IS NOT NULL;
		''', conn)
		# print(len(check_E8))
		check_I8 = pd.read_sql_query(f'''
		SELECT DISTINCT a.PERIOD, b.PERIOD
		FROM {PRE_FINAL_TABLE_IMP} a
		LEFT JOIN [SRC_Macao].[dbo].[I8] b
		ON a.PERIOD = b.PERIOD
		WHERE b.PERIOD IS NOT NULL;
		''', conn)
		# print(len(check_I8))

		insrtCondition = False
		if list(genericChecks.all())[0] and len(check_E8) == 0 and len(check_I8) == 0:
			# INSERT INTO ARCHIVE
			cursor.execute("""
			INSERT INTO [SRC_Macao].[dbo].[SRC_EXP_IMP_ARCHIVE]
			SELECT [FILENAME]
				,[FILE_YEAR]
				,[FILE_MONTH]
				,[FLOW]
				,[YEAR]
				,[MONTH]
				,[COUNTRY]
				,[COUNTRY_CODE]
				,[NCEM/HS code]
				,[Commodity description]
				,CAST([Quantity] AS DECIMAL) AS [Quantity]
				,[Unit]
				,CAST([Weight (KG)] AS DECIMAL) AS [Weight (KG)]
				,CAST([Macao Pataca] AS DECIMAL) AS [Macao Pataca]
			FROM  [SRC_Macao].[dbo].[TEMP];
			""") 
			cursor.commit()

			# insert into final tables
			cursor.execute('''
			-- INSERT EXPORTS
			DECLARE @MinPeriod VARCHAR(6), @MaxPeriod VARCHAR(6);
			SELECT @MinPeriod = MIN(PERIOD), @MaxPeriod = MAX(PERIOD)
			FROM [SRC_Macao].[dbo].[TEMP_STEP5_EXP];

			DELETE FROM [SRC_Macao].[dbo].[E8] 
			WHERE PERIOD BETWEEN @MinPeriod AND @MaxPeriod;

			INSERT INTO [SRC_Macao].[dbo].[E8]
			SELECT [CTY_RPT]
			,[CTY_PTN]
			,[COMMODITY]
			,[PERIOD]
			,[YR]
			,[MON]
			,[VALUE_USD] AS VALUE
			,[VALUE_MOP]
			,[UNIT1]
			,[QTY1]
			,[UNIT2]
			,[QTY2]
			FROM [SRC_Macao].[dbo].[TEMP_STEP5_EXP];

			-- INSERT IMPORTS
			SELECT @MinPeriod = MIN(PERIOD), @MaxPeriod = MAX(PERIOD)
			FROM [SRC_Macao].[dbo].[TEMP_STEP5_IMP];

			DELETE FROM [SRC_Macao].[dbo].[I8] 
			WHERE PERIOD BETWEEN @MinPeriod AND @MaxPeriod;

			INSERT INTO [SRC_Macao].[dbo].[I8]
			SELECT [CTY_RPT]
			,[CTY_PTN]
			,[COMMODITY]
			,[PERIOD]
			,[YR]
			,[MON]
			,[VALUE_USD] AS VALUE
			,[VALUE_MOP]
			,[UNIT1]
			,[QTY1]
			,[UNIT2]
			,[QTY2]
			FROM [SRC_Macao].[dbo].[TEMP_STEP5_IMP];
			
			''')
			cursor.commit()

			# drop tables
			cursor.execute('''
			DROP TABLE [SRC_Macao].[dbo].[TEMP_SUPPRESSED];
			DROP TABLE [SRC_Macao].[dbo].[TEMP_SUPPRESSED_ALL_COUNTRIES];
			DROP TABLE [SRC_Macao].[dbo].[TEMP_ALL_COUNTRIES];
			DROP TABLE [SRC_Macao].[dbo].[TEMP_TOTALS];
			DROP TABLE [SRC_Macao].[dbo].[TEMP_TOTALS_STEP1];
			--DROP TABLE [SRC_Macao].[dbo].[TEMP];
			DROP TABLE [SRC_Macao].[dbo].[TEMP_STEP1];
			DROP TABLE [SRC_Macao].[dbo].[TEMP_STEP2];
			DROP TABLE [SRC_Macao].[dbo].[TEMP_STEP3_EXP];
			DROP TABLE [SRC_Macao].[dbo].[TEMP_STEP3_IMP];
			DROP TABLE [SRC_Macao].[dbo].[TEMP_STEP4_EXP];
			DROP TABLE [SRC_Macao].[dbo].[TEMP_STEP4_IMP];
			--DROP TABLE [SRC_Macao].[dbo].[TEMP_STEP5_EXP];
			--DROP TABLE [SRC_Macao].[dbo].[TEMP_STEP5_IMP];
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
					,"y.zeng@tradedatamonitor.com"
					,"a.chan@tradedatamonitor.com"
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

	mo = ETL('Macao')
	mo.run()
