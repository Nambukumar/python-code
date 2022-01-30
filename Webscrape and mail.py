import os
import unittest
import smtplib
import requests
import warnings
import time
import numpy as np
import datetime
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from cssselect import GenericTranslator, SelectorError
from bs4 import BeautifulSoup
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.message import EmailMessage

#os.system("taskkill /f /im chrome.exe")

x = datetime.datetime.now()
day = x.day
month = x.strftime("%B")
toshi = x.year
gap = " "
date = str(day) + gap + str(month) + gap + str(toshi)

#DB_USER = Create Environment variable for your gmail ID
#DB_PASS = Create Environment variable for your gmail password

EMAIL_ADDRESS = os.environ.get('DB_USER')
EMAIL_PASSWORD = os.environ.get('DB_PASS')

recipients = ['sample1@samplecom','sample2@sample.com' ]

url = 'your URL goes here'

browser = webdriver.Chrome(executable_path=ChromeDriverManager().install())
browser.get(url)
time.sleep(10)

warnings.filterwarnings("ignore", category=DeprecationWarning) 
button = browser.find_element_by_css_selector("[tag = value]")
button.click()
time.sleep(10)	
certificate = browser.find_element_by_css_selector("[tag = value]")
#print ('Certifications :', certificate.text)
split_count = np.array(certificate.text)

msg = EmailMessage()
msg['Subject'] =  str(date) + gap + '[CERTIFICATION Count]'
msg['From'] = EMAIL_ADDRESS
msg['To'] = recipients
msg.set_content(certificate.text)

with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
		smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
		smtp.send_message(msg) 

browser.close()
#os.system("taskkill /f /im chrome.exe")