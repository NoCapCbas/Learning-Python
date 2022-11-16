# Approx. 50 files in one hour without being blocked.
# Approx. 5 consecutive hours until being blocked.
import random 
import string
import pyautogui
import logging
from datetime import datetime
from datetime import timedelta
import pyodbc
from time import sleep
import subprocess
import pandas as pd
import os 
import glob
import sys
import math
from pywinauto.application import Application
# selenium imports
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui  import Select

# create logger
logger = logging.getLogger('Log')
logger.setLevel(logging.INFO)
# create file handler and set level
handler = logging.FileHandler(filename=f"Y:\\_PyScripts\\Damon\\Log\\China\\ChinaMadeEZ.log", mode='w')
handler.setLevel(logging.INFO)
# create formatter
format = logging.Formatter('%(asctime)s %(levelname)s:%(message)s', datefmt='%b-%d-%Y %H:%M:%S')
# add formatter to handler
handler.setFormatter(format)
# add handler to logger
logger.addHandler(handler)
# console logger
consoleHandler = logging.StreamHandler()
logger.addHandler(consoleHandler)

# edit variables
periodsToScrape = [
	'202101',
	'202102',
	'202103',
	'202104',
	'202105',
	'202106',
	'202107',
	'202108',
	'202109',
	'202110',
]
fieldsToSelect = {
	'outerField1':'CODE_TS',
	'outerField2':'ORIGIN_COUNTRY',
	'outerField3':'TRADE_MODE',
	'outerField4':'TRADE_CO_PORT',
}
FLOWS = [
	{'code': 0, 'name': 'EXP_TOTAL'},
	{'code': 1, 'name': 'IMP_TOTAL'},
	{'code': 0, 'name': 'EXP'},
	{'code': 1, 'name': 'IMP'},
]
rows_per_page = 10000
# end edit variables

# URLS
HOME_URL = 'http://43.248.49.97/'

# PATHS
basePath = f"Y:\\_PyScripts\\Damon\\Log\\"
imgPath = "Y:\\_PyScripts\\Damon\\China\\"
savePath = basePath + f"China\\DUMP\\"
if not os.path.exists(basePath):
	os.mkdir(basePath)
if not os.path.exists(savePath):
	os.mkdir(savePath)

def findImage(imageString, confidence=0.9, grayscale=False):
	image = None
	# current time
	startNow = datetime.now()
	future_time = startNow + timedelta(minutes=4)
	logger.info(f'\t\tSearching for {imageString}...  Now:{startNow} Max Limit:{future_time}')
	while image == None:
		now = datetime.now()
		image = pyautogui.locateOnScreen(imageString, confidence=confidence, grayscale=True)
		pageError = pyautogui.locateOnScreen(f'{imgPath}pageNotWorking.png')
		if now >= future_time:
			raise ValueError(f'Image Load Time Limit of 4 minutes reached: {future_time}')
		if pageError and image == None and now >= startNow + timedelta(minutes=1): 
			raise ValueError(f'Page did not Load.')
	imageCenter = pyautogui.center(image)
	print(f'\t\t\t{imageString} found.')
	return image, imageCenter

def triggerHotSpotVPN():
	screenSizeX, screenSizeY = pyautogui.size()
	screenSizeY = 200 
	# click to distract finder if open
	pyautogui.click(screenSizeX, screenSizeY)
	Locations = [
		# Working
		'China',
		'Russia',
		'Hong Kong','Los Angeles',
		'Azerbaijan','Colombia','Portland',
		'Argentina','Washington DC','Bahamas','Adelaide',
		'Greece', 'Bangladesh', 'Kyrgyzstan','New Jersey',
		'Kazakhstan', 'Atlanta','Boston','Indianapolis',
		'Kansas City','Miami','Perth','Algeria','Armenia','St. Louis'
		# Not Sure
		'Seattle','Bosnia & Herzegovina','India',
		'Belarus', 'Belgium', 'Belize', 'Bhutan', 'Brunei',
		'United Kingdom', 'Coventry','Russia','Georgia',
		'Barcelona', 'Sydney','Czech Republic','Iceland',
		'Denmark','Costa Rica','Ecuador','Finland',
		'Croatia','Austria', 'Belgium', 'Brazil','Hungary',
		'Germany', 'Ireland', 'Italy', 'Rome', 'Ukraine',
		'Bulgaria','Cambodia','Chile','Egypt','Estonia',
		'Indonesia','Isle of Man','Israel','Japan','Laos',
		'Latvia','Liechtenstein','Lithuania','Luxembourg',
		'Malaysia','Malta','Mexico','Moldova','Monaco',
		'Montenegro','Nepal','Netherlands','New Zealand',
		'Norway','Pakistan','Panama','Peru','Philippines',
		'Poland','Portugal','Singapore','Slovakia','South Africa',
		'South Korea','Sweden','Switzerland','Taiwan','Thailand',
		'Turkey','United Arab Emirates','Uruguay','Venezuela',
		'Vietnam',
		# Not Working
		'Melbourne','Milan','Chicago','Denver','Orlando',
		'Columbus','New York','Charlotte','Dallas','Houston',
		'San Jose','Brisbane','Las Vegas',
		'Philadelphia','Phoenix','Portland','San Francisco'
	]
	# open hotspot shield
	print('Searching for Hotspot Shield...')
	pathToHotspot = subprocess.check_output(['where', '/R', 'C:\\', '*hsscp.exe'])
	pathToHotspot = str(pathToHotspot.decode('utf-8').replace('\\', '\\\\')).strip()
	if os.path.exits(pathToHotspot):
		print('Opening Hotspot Shield...')
		app = Application().start(pathToHotspot)
	else:
		raise ValueError('Hotspot Shield Not Found.')
	# END open hotspotshield

	# searching if hotspotsheild was open 
	temp, tempCenter = findImage(f'{imgPath}homeHotSpot.png')
	pyautogui.click(tempCenter.x, tempCenter.y)
	sleep(3)

	connectToHotspot = pyautogui.locateOnScreen(f'{imgPath}connectToHotspot.png')
	disconnectHotSpot = pyautogui.locateOnScreen(f'{imgPath}disconnectHotSpot.png')
	
	sleep(3)
	if connectToHotspot:
		pass
	else: 
		disconnectHotSpotCenter = pyautogui.center(disconnectHotSpot)
		pyautogui.click(disconnectHotSpotCenter.x, disconnectHotSpotCenter.y)
		sleep(3)
	temp, tempCenter = findImage(f'{imgPath}virtualLocation.png')
	pyautogui.click(tempCenter.x, tempCenter.y)

	randLocation = random.choice(Locations)
	pyautogui.write(randLocation)
	sleep(3)
	# if looping through more than one cty
	if pyautogui.locateOnScreen(f'{imgPath}ctyCheck.png'):
		triggerHotSpotVPN()
	else:
		temp, tempCenter = findImage(f'{imgPath}hotspot_connect_to_cty.png')
		pyautogui.click(tempCenter.x + 520, tempCenter.y)

		findImage(f'{imgPath}disconnectHotSpot.png')
		print('\t\t\tHotspot On.')

def scrambleChrome(pathToChromedriver):
	# opens chromedriver in read byte mode
	with open(pathToChromedriver, 'rb+') as f: 
		# reads chromedriver 
		file = f.readlines()
		# picks 3 random letters
		letter1 = random.choice(string.ascii_letters)
		letter2 = random.choice(string.ascii_letters)
		letter3 = random.choice(string.ascii_letters)
		count = 0
		# finds and replaces line
		for line in file: 
			if "var key = '$cdc_asdjflasutopfhvcZLmcfl_';" in line.decode('ANSI'):
				currentVal = line.decode('ANSI')
				newVal = currentVal.replace('$cdc_asd', f'${letter1}{letter2}{letter3}_asd')
				print(currentVal)
				file[count] = newVal.encode('latin1')
				print(file[count])
			count += 1
	# opens chromedriver in write byte mode
	with open(pathToChromedriver, 'wb') as f: 
		# writes to file
		for line in file:
			f.write(bytes(line))

def connect(URL, instance = 0):
	driver = 0
	try: 
		print('Connecting...')
		# window instance
		instance += 1
		print('Instance: ' + str(instance))
		
		# triggerHotSpotVPN()

		chrome_options = webdriver.ChromeOptions()
		# disables flag to visit website undetected, Bypass cloudflare detection
		chrome_options.add_argument("--disable-blink-features=AutomationControlled")
		chrome_options.add_experimental_option("useAutomationExtension", False)
		chrome_options.add_experimental_option("excludeSwitches",["enable-automation"])
		chrome_options.add_argument("window-size=1775,950")
	
		prefs = {'download.default_directory': savePath, 
		'profile.default_content_setting_values.automatic_downloads': 'Hotspot Shield'
		}
		chrome_options.add_experimental_option('prefs', prefs)
		executable_path = ChromeDriverManager().install()
		# executable_path = "C:\\Users\\admin9\\Documents\\chromedriver.exe"
		scrambleChrome(executable_path)
		driver = webdriver.Chrome(executable_path=executable_path, chrome_options=chrome_options)
		driver.get(URL)
		sleep(3)
	except Exception as e:
		logger.exception(e)
		# sleep(10)
		if driver != 0:  
			driver.quit()
		driver = connect(URL, instance)
	
	return driver  
	

def main():
	logger.info('Launching...')
	scriptOver = False
	driver = 0
	restart_browser = True
	while scriptOver == False:
		if restart_browser:
			logger.info('RESTARTING BROWSER...')
			driver = connect(HOME_URL)
			restart_browser = True
		

		logger.info('\t\tBeginnning Downloads...')
		sleep(5)
		try: 
			scriptOver = False
			
			for period in periodsToScrape:
				for flow in FLOWS:

					driver.switch_to.default_content()
					try: 
						logger.info('Grabbing iframe')
						# //*[@id="iframe_box"]/div[1]/iframe
						xpath = '//*[@id="iframe_box"]/div[2]/iframe'
						iframe = WebDriverWait(driver, 240).until(EC.presence_of_element_located((By.XPATH, xpath)))
						logger.info('Switching iframe')
						driver.switch_to.frame(iframe)
						logger.info('\tSelecting Year...')
						Select(WebDriverWait(driver, 120).until(EC.element_to_be_clickable(((By.ID, 'year'))))).select_by_value(period[:-2])					
					except Exception as e:
						driver.switch_to.default_content()
						logger.info('Grabbing iframe')
						xpath = '//*[@id="iframe_box"]/div[1]/iframe'
						iframe = WebDriverWait(driver, 240).until(EC.presence_of_element_located((By.XPATH, xpath)))
						logger.info('Switching iframe')
						driver.switch_to.frame(iframe)
						logger.info('\tSelecting Year...')
						Select(WebDriverWait(driver, 120).until(EC.element_to_be_clickable(((By.ID, 'year'))))).select_by_value(period[:-2])

					sleep(5)
					logger.info('\tSelecting Start Month...')
					WebDriverWait(driver, 45).until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="startMonth"]/option[{str(int(period[-2:]))}]'))).click()
					sleep(3)
					logger.info('\tSelecting End Month...')
					WebDriverWait(driver, 45).until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="endMonth"]/option[{str(int(period[-2:]))}]'))).click()
					sleep(3)

					iframeNUM = 1
					file_name = savePath + f'{period}{flow["name"]}'
					if os.path.exists(f'{file_name}_final.csv'):
						logger.info(f'\t\t\tFinal file exists for {period}{flow["name"]}...')
						continue
					logger.info(f'\t\t\t {period} {flow["name"]}')

					logger.info(f'\t\t\t\tSelecting {flow["name"]}...')
					driver.execute_script(f"$('input[name=iEType][value={flow['code']}]').click();")
					
					logger.info(f'\t\t\t\tSelecting USD...')
					WebDriverWait(driver, 45).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Search_form"]/table/tbody/tr[1]/td[2]/div[3]/input[2]'))).click()

					if 'outerField1' in fieldsToSelect.keys():
						logger.info('\t\t\t\tSelecting Field 1...')
						WebDriverWait(driver, 45).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="outerField1"]/option[2]'))).click()
					
					if 'outerField2' in fieldsToSelect.keys() and 'TOTAL' not in flow['name']:
						logger.info('\t\t\t\tSelecting Field 2...')
						WebDriverWait(driver, 45).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="outerField2"]/option[3]'))).click()

					if 'outerField3' in fieldsToSelect.keys() and 'TOTAL' not in flow['name']:
						logger.info('\t\t\t\tSelecting Field 3...')
						WebDriverWait(driver, 45).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="outerField3"]/option[4]'))).click()

					if 'outerField4' in fieldsToSelect.keys() and 'TOTAL' not in flow['name']:
						logger.info('\t\t\t\tSelecting Field 4...')
						WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="outerField4"]/option[5]'))).click()
		
					logger.info(f'\t\t\t\tSubmiting form...')
					WebDriverWait(driver, 45).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="doSearch"]'))).click()
					

					# check if pop up is recaptcha or confirmation of data
					logger.info(f'\t\t\t\tBeating ReCaptcha...iFrame:{iframeNUM}')
					try:
						captchaFrame = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, f'//*[@id="layui-layer-iframe{iframeNUM}"]')))
					except:
						logger.info(f'\t\t\t\t\tConfirming old period...')
						driver.execute_script("$('.layui-layer-btn0').click()")
						iframeNUM += 1
						logger.info(f'\t\t\t\tRetrying ReCaptcha...iFrame:{iframeNUM}')
						captchaFrame = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, f'//*[@id="layui-layer-iframe{iframeNUM}"]')))
						iframeNUM += 1

					logger.info(f'\t\t\t\tSwitching to captchaFrame...')
					driver.switch_to.frame(captchaFrame)
					logger.info(f'\t\t\t\tSlider add active...')
					driver.execute_script("$('.sliderContainer').addClass('sliderContainer_active');")
					WebDriverWait(driver, 120).until(EC.presence_of_element_located(((By.CSS_SELECTOR, '.sliderContainer.sliderContainer_active'))))
					logger.info(f'\t\t\t\tSlider to success...')
					driver.execute_script("$('.sliderContainer.sliderContainer_active').addClass('sliderContainer_success');")
					WebDriverWait(driver, 120).until(EC.presence_of_element_located(((By.CSS_SELECTOR, '.sliderContainer.sliderContainer_active.sliderContainer_success'))))
					logger.info(f'\t\t\t\tSlider remove active')
					driver.execute_script("$('.sliderContainer.sliderContainer_active.sliderContainer_success').removeClass('sliderContainer_active');")
					WebDriverWait(driver, 120).until(EC.presence_of_element_located(((By.CSS_SELECTOR, '.sliderContainer.sliderContainer_success'))))
					
					logger.info('\t\t\tGo to Data Page...')
					driver.execute_script('''
					var serialize = $("#Search_form").serialize();
					parent.window.location.href="/queryData/queryDataList?"+serialize; 
					''')
					iframe = WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, xpath)))
					driver.switch_to.frame(iframe)

					logger.info(f'\t\t\tGrabbing total rows...')
					totalRows = WebDriverWait(driver, 120).until(EC.presence_of_element_located(((By.ID, 'totalSize')))).get_attribute("value")
					
					logger.info(f'\t\t\tGrabbing total pages...')
					totalPages = math.ceil(int(totalRows)/rows_per_page)
					sleep(5)				

					logger.info(f'\t\t\tBeginning page...')
					MASTER = pd.DataFrame()
					for page in range(1, int(totalPages)+1):
						logger.info(f'\t\t\t\tPage {page}...')
						if page == int(totalPages):
							uniqFILENAME = file_name + f'_final.csv'
						else:
							uniqFILENAME = file_name + f'_{page}.csv'
						if os.path.exists(f'{uniqFILENAME}'):
							tempDF = pd.read_csv(uniqFILENAME)
							MASTER = pd.concat([MASTER, tempDF], ignore_index=True)
							logger.info(f'\t\t\t{uniqFILENAME} file exists for {period}{flow["name"]}...')
							continue
						# add total amount of rows of data per page option
						driver.execute_script(f"$('#pageSize').append(new Option('{rows_per_page}', '{rows_per_page}'));")
						sleep(3)
						# edit onchange
						driver.execute_script(f"$('#pageSize').attr('onChange','doSearch({page})');")
						sleep(3)
						# select total amount of rows of data per page option and moves to next page
						driver.execute_script(f"$('#pageSize').val('{rows_per_page}').trigger('change');")
						sleep(8)
						pageLOADING = True
						# current time
						now = datetime.now()
						future_time = now + timedelta(minutes=5)
						logger.info(f'\t\t\t\t\tLoading Page...  Now:{now} Max Limit:{future_time}')
						while pageLOADING:
							now = datetime.now()
							# wait for page to load
							WebDriverWait(driver, 180).until(EC.presence_of_element_located(((By.ID, 'table'))))
							sleep(5)
							# verify data is on screen and grab table
							WebDriverWait(driver,120).until(EC.presence_of_element_located(((By.ID, 'totalPages'))))
							htmlTABLE = WebDriverWait(driver, 120).until(EC.presence_of_element_located(((By.ID, 'table')))).get_attribute('outerHTML')
							src_page_num = WebDriverWait(driver, 120).until(EC.presence_of_element_located(((By.ID, 'change-page-num')))).get_attribute('value')
							# turn HTML table to dataframe and turn to csv
							tempDF = pd.read_html(htmlTABLE)[0]
							# logger.info(len(tempDF))
							# logger.info(tempDF)
							
							if (len(tempDF) == rows_per_page or (uniqFILENAME.split('_')[-1] == 'final.csv' and len(tempDF) == int(totalRows)%rows_per_page)) and str(src_page_num) == str(page):
								# print(MASTER)
								# print(tempDF)
								if len(MASTER) != 0:
									merged = pd.merge(MASTER, tempDF, on=list(MASTER.columns), how='outer', indicator=True)
									# print(merged)
									if 'both' in list(merged[merged.columns[-1]]):
										# check if able to return to home
										driver.switch_to.default_content()
										driver.execute_script("$('#aaa').click();")
										restart_browser = False
										logger.info(f'PAGE ERROR: Data Collected Exists in MASTER. RESTART.')
										raise ValueError(f'PAGE ERROR: Data Collected Exists in MASTER. RESTART.')

								logger.info(f'\t\t\tDumping {uniqFILENAME}...')
								MASTER = pd.concat([MASTER, tempDF], ignore_index=True)
								tempDF.to_csv(uniqFILENAME, index=False)
								pageLOADING = False
								sleep(8)
							if now >= future_time:
								# Try to get back to home page
								driver.switch_to.default_content()
								driver.execute_script("$('#aaa').click();")
								restart_browser = False
								raise ValueError(f'Download Time Limit of 5 minutes reached: {future_time}')
					MASTER.to_csv(savePath + f'ALL_{period}{flow["name"]}', index=False)
					# return to main menu
					driver.switch_to.default_content()
					driver.execute_script("$('#aaa').click();")
					

			scriptOver = True
		except Exception as e:
			logger.exception(e)
			if driver == 0 or restart_browser == False:
				pass
			else:
				driver.quit()
			logger.info(f'***Error***: {e}')
			if scriptOver:
				sys.exit()


if __name__ == '__main__':
	main()