# sql server import
from stat import filemode
import pyodbc
# pandas import
import pandas as pd
# datetime import
from datetime import datetime
from dateutil.relativedelta import relativedelta
# email import
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
# misc import
import shutil
import glob
import os
from six.moves import urllib
import requests
from time import sleep
import json
import time
import win32com.client as win32
from pywintypes import com_error
from pathlib import Path
import sys
import numpy as np
import logging
import tkinter as tk
# create logger
logger = logging.getLogger('Log')
logger.setLevel(logging.INFO)
# create file handler and set level
handler = logging.FileHandler(filename='practice.log', mode='w')
handler.setLevel(logging.INFO)
# create formatter
format = logging.Formatter('%(asctime)s %(levelname)s:%(message)s', datefmt='%b-%d-%Y %H:%M:%S')
# add formatter to handler
handler.setFormatter(format)
# add handler to logger
logger.addHandler(handler)

logging.debug('This is debug')
logging.info('This is info')
logging.warning('This is a warning')
logging.error('This is an error')
logging.critical('This is critical')

logger.debug('This is debug')
logger.info('This is info')
logger.warning('This is a warning')
logger.error('This is an error')
logger.critical('This is critical')
