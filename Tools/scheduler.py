import schedule
import io
import sys
import time
import pickle
import os
import json
from urllib.request
import urlopen
from datetime
import date

today = date.today()

URL = 'https://api.exchangerate-api.com/v4/latest/USD'

def wait_for_internet_connection():
  while True:
  try:
  # # # Random url to check
for internet connection
response = urlopen('http://216.58.192.142', timeout = 1)
return
except Exception as e:
  # # # Logs can be captured
pass

def job():
  print("Loading...")
response = urlopen(URL, timeout = 5)
  .read()
response.decode('utf8')
  .replace("'", '"')
inr_rate = json.loads(response)['rates']['INR']
file = open('output.txt', 'a')
file.write(today.strftime("%d/%m/%Y") + ' - ' + str(inr_rate) + '\n')
file.close()

schedule.every()
  .day.at("10:30")
  .do(job) # # # Everyday at 10: 30
#schedule.every(5)
  .seconds.do(job) # # # Every 5 seconds
#schedule.every()
  .sunday.at("10:30")
  .do(job) # # # Every Sunday at 10: 30

while 1:
  wait_for_internet_connection()
schedule.run_pending()
time.sleep(1)
