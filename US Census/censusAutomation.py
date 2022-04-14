from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import requests
import shutil
import os
import glob
import pandas as pd
from datetime import datetime
from datetime import timedelta
#email imports
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

list = []
pathToTotal = r"Y:\_PyScripts\Damon\Log\usCensus\totalCENSUS.txt"
pathToLog = r"Y:\_PyScripts\Damon\Log\usCensus\downloadLog.txt"
listDiff = []
downloadPath = "Y:\\_PyScripts\\Damon\\Log\\usCensus\\Downloads"

def connect(): 
    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory': "Y:\\_PyScripts\\Damon\\Log\\usCensus\\Downloads"}
    chrome_options.add_experimental_option('prefs', prefs)
    URL = 'https://usatrade.census.gov/index.php'
    driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=chrome_options)
    driver.get(URL)
    return driver

def usCensus(driver):
    # clear download folder of previous downloads
    print('Clearing Downloads Folder...')
    for file in os.listdir(downloadPath): 
        os.remove(os.path.join(downloadPath, file))
    print('Download Folder Cleared.')
    print('Grabbing List of Data to scrape...')
    #grab the Total list
    df = pd.read_csv(pathToTotal)
    first_column = df.iloc[:, 0]
    totalList = first_column.tolist()
    #grab the download log list
    df = pd.read_csv(pathToLog)
    first_column = df.iloc[:, 0]
    downloadList = first_column.tolist()

    #get the difference between the two lists
    for item in totalList:
        if item not in downloadList:
            print(item)
            listDiff.append(item)
    #print(listDiff)
    def create_dir(dirName): #pass in name of dir you want to create
        #Create new dir
        path = "Y:\\_PyScripts\\Damon\\Log\\usCensus\\"  # path to create new dir
        new_dir = str(dirName)  # dir name that files will save to
        path = path + new_dir
        try:
            os.mkdir(path)
        except:
            pass
        print(f'Directory {new_dir} created.')
        return new_dir
    
    dir = create_dir('CensusEXPORTS')
    

    # clicks login button using css selector
    sleep(2)
    driver.find_element_by_css_selector("#userloginbutton").click()


    # Logins to website Using Credentials
    #User: 8YRN87M
    #Pass: Newtradedata1!
    sleep(3)
    userId = driver.find_element_by_name("struserid")
    userId.send_keys('8YRN87M')

    password = driver.find_element_by_name("pwdfld")
    password.send_keys('TradeData2257$')

    submit = driver.find_element_by_name("submit")
    submit.click()

    #Clicks on State export or import
    #Export: 6541
    #Import: 13307
    sleep(3)
    l = driver.find_element_by_xpath('//a[@href="javascript:OnLoadFirstDimension(6541)"]')
    l.click()
    # Selects Measures (ALL)
    sleep(3)
    l=driver.find_element_by_xpath("//a[@title='Select members of Measures']")
    l.click()

    sleep(3)
    l=driver.find_elements_by_xpath("//input[@type='checkbox']")
    for checkbox in l:
        checkbox.click()
    # Selects Countries (ALL)
    l=driver.find_element_by_xpath("//a[@title='Select members of Country']")
    l.click()
    sleep(3)

    l=driver.find_element_by_xpath('//*[@id="ctl00_MainContent_RadMembersTree"]/ul/li/div/label/input[2]')
    l.click()

    l=driver.find_element_by_xpath('//*[@id="ctl00_MainContent_RadMembersTree"]/ul/li/div/label/input[1]')
    l.click()

    #use the difference list as variables to use to get the rest of the data
    listDiff.reverse() # makes it start from 2007 - 2002
    for item in listDiff:
        itemList = item.split('.')
        i = itemList[0] # year
        j = itemList[1] # month
        k = itemList[2] # state
        x = itemList[3] # commodity
        xpath_time_year_expandYear = f'//*[@id="ctl00_MainContent_RadMembersTree"]/ul/li[{str(i)}]/div/span[2]'
        xpath_time_year_minimizeYear = f'//*[@id="ctl00_MainContent_RadMembersTree"]/ul/li[{str(i)}]/div/span[1]'

        xpath_time_year_expandYear = f'//*[@id="ctl00_MainContent_RadMembersTree"]/ul/li[{str(i)}]/div/span[2]'
        xpath_time_year_minimizeYear = f'//*[@id="ctl00_MainContent_RadMembersTree"]/ul/li[{str(i)}]/div/span[1]'

        # Selects Time
        #print('Selecting Time Tab')
        try:
            l=driver.find_element_by_xpath("//a[@title='Select members of Time - time series hierarchy']")
            l.click()
            sleep(3)
        except:
            pass

        xpath_year_text = driver.find_element_by_xpath(f'//*[@id="ctl00_MainContent_RadMembersTree"]/ul/li[{str(i)}]/div/label/span').text
        yearDir = create_dir(f'{dir}\\{xpath_year_text}')

        try:
            l=driver.find_element_by_xpath(xpath_time_year_expandYear)
            l.click()
            sleep(3)
        except:
            pass

            # Loops through Time using xpath positioning
            # Edit Scrape (months) 1-12

        xpath_time_month = f'//*[@id="ctl00_MainContent_RadMembersTree"]/ul/li[{str(i)}]/ul/li[{str(j)}]/div/label/input'
        xpath_Month_text = driver.find_element_by_xpath(f'//*[@id="ctl00_MainContent_RadMembersTree"]/ul/li[{str(i)}]/ul/li[{str(j)}]/div/label/span').text
        monthDir = create_dir(f'{yearDir}\\{xpath_Month_text}')

        # checks month
        l = driver.find_element_by_xpath(xpath_time_month)
        l.click()
        sleep(2)

        # Selects State Category
        try:
            l=driver.find_element_by_xpath("//a[@title='Select members of State']")
            l.click()
            sleep(3)
        except:
            pass

                # Loops through States using xpath positioning
                # Edit Scrape (States) 1-54

        xpath_State = '//*[@id="ctl00_MainContent_RadMembersTree"]/ul/li/ul/li[' + str(k) + ']/div/label/input'
        # Selects State Category
        try:
            l=driver.find_element_by_xpath("//a[@title='Select members of State']")
            l.click()
            sleep(3)
        except:
            pass
        xpath_State_text = driver.find_element_by_xpath(f'//*[@id="ctl00_MainContent_RadMembersTree"]/ul/li/ul/li[{str(k)}]/div/label/span').text
        stateDir = create_dir(f'{monthDir}\\{xpath_State_text}')
        # check state box
        l = driver.find_element_by_xpath(xpath_State)
        l.click()
        sleep(2)


                    # Loops through commodities using xpath positioning
                    # Edit Srape (Chapters) 1-97

        xpath_commodity_check = '//*[@id="ctl00_MainContent_RadMembersTree"]/ul/li/ul/li[' + str(x) + ']/div/label/input[2]'
        xpath_commodity_uncheck = '//*[@id="ctl00_MainContent_RadMembersTree"]/ul/li/ul/li[' + str(x) + ']/div/label/input[3]'
        xpath_commodity_minimize = '//*[@id="ctl00_MainContent_RadMembersTree"]/ul/li/ul/li[' + str(x) + ']/div/span[2]'


        # Selects Commodity tab
        try:
            l=driver.find_element_by_xpath("//a[@title='Select members of Commodity']")
            l.click()
            sleep(3)
        except:
            pass
        xpath_chapter_text = driver.find_element_by_xpath(f'//*[@id="ctl00_MainContent_RadMembersTree"]/ul/li/ul/li[{str(x)}]/div/label/span').text
        # check commodity box
        l = driver.find_element_by_xpath(xpath_commodity_check)
        l.click()
        sleep(3)

    ##################################################################################################################################################################
        #Download file
        l = driver.find_element_by_xpath('//*[@id="breadcrumbs"]/font[3]/a')
        l.click()
        sleep(3)

        l = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_RadContentToolBar"]/div/div/div/ul/li[9]/a/span/span/span/img')
        l.click()
        sleep(3)
        # Download Format: Select Comma delimited
        l = driver.find_element_by_xpath('//*[@id="DownloadId"]/option[2]')
        l.click()
        sleep(3)
        # Begin Download
        l = driver.find_element_by_xpath('//*[@id="OptionDialog"]/table/tbody/tr[8]/td/input[1]')
        l.click()
        sleep(5)
        #file management
        condition = False
        print('\tSearching for downloaded file...')
        startTime = datetime.now()
        raiseErrorTimer = startTime + timedelta(minutes=10)
        while condition == False:
            if datetime.now() >= raiseErrorTimer: 
                raise Exception('raiseErrorTimer Triggered. Searching for download took longer than 10 minutes.')

            try:

                list_of_files = glob.glob('Y:\\_PyScripts\\Damon\\Log\\usCensus\\Downloads\\State Exports by HS Commodities.csv') # * means all if need specific format then *.csv
                latest_file = max(list_of_files, key=os.path.getctime)
                #print(f'--{latest_file}')

                if latest_file == 'Y:\\_PyScripts\\Damon\\Log\\usCensus\\Downloads\\State Exports by HS Commodities.csv':

                    condition = True

            except:
                sleep(3)
                pass

        #renaming file
        new_file = f'{xpath_chapter_text[:2]}_{xpath_State_text}_{xpath_Month_text}.csv'
        os.rename(latest_file,f'Y:\\_PyScripts\\Damon\\Log\\usCensus\\Downloads\\{new_file}')
        #moving File
        shutil.move(f'Y:\\_PyScripts\\Damon\\Log\\usCensus\\Downloads\\{new_file}', f'Y:\\_PyScripts\\Damon\\Log\\usCensus\\{stateDir}\\{new_file}')
        sleep(1)
        print('-----------------------------------')
        print(f'{stateDir}\\{new_file} Downloaded')
        print('-----------------------------------')
        #UPDATES log
        with open(pathToLog, 'a') as f:
                f.write(f"{i}.{j}.{k}.{x}\n")
    ##################################################################################################################################################################




        # Selects Commodity
        try:
            l=driver.find_element_by_xpath("//a[@title='Select members of Commodity']")
            l.click()
            sleep(3)
        except:
            pass

        # uncheck commodity box
        l = driver.find_element_by_xpath(xpath_commodity_uncheck)
        l.click()
        sleep(3)
        print(f'Finished Chapter {xpath_chapter_text[:2]}, MONTH {j} STATE {k} ')

        # minimize finished commodity box
        l = driver.find_element_by_xpath(xpath_commodity_minimize)
        l.click()
        sleep(3)





        # Selects State Category
        try:
            l=driver.find_element_by_xpath("//a[@title='Select members of State']")
            l.click()
            sleep(3)
        except:
            pass
        # uncheck state box
        l = driver.find_element_by_xpath(xpath_State)
        l.click()
        sleep(2)
        print(f'Finished State {xpath_State_text}')



        #Selects Time Tab
        try:
            l=driver.find_element_by_xpath("//a[@title='Select members of Time - time series hierarchy']")
            l.click()
            sleep(3)
        except:
            pass

        # unchecks month
        l = driver.find_element_by_xpath(xpath_time_month)
        l.click()
        sleep(2)
        #closes year of inner loop that just ended
        try:
            l=driver.find_element_by_xpath(xpath_time_year_expandYear)
            l.click()
            sleep(3)
        except:
            pass

    driver.close()

def EmailErrorCheck(e):
    print('Sending Email...')
    port = 465
    smtp_server = "smtp.gmail.com"
    sender_email = "tdmUsageAlert@gmail.com"
    recipients = ["DDiaz@tradedatamonitor.com",]
    password = "tdm12345$"     #tdm12345$
    message = MIMEMultipart("alternative")
    message["Subject"] = 'US Census Automation Error Check'
    message["from"] = "TDM Data Team"
    message["To"] = ", ".join(recipients)
    html = f"""\
    <html>
        <body>
            <p>{e}</p>
        </body>
      </html>
      """
    content = MIMEText(html, "html")
    message.attach(content)
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtp_server,port, context = context) as server:
        server.login("tdmUsageAlert@gmail.com", password)
        server.sendmail(sender_email, recipients, message.as_string())


def recursion():
    driver = connect()
    try:
        usCensus(driver)
    except Exception as e:
        
        driver.close()
        
        print('*******************************************')
        print(e)
        print('*******************************************')
        EmailErrorCheck(e)
        sleep(10)
        recursion()

recursion()
