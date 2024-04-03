from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


import win32com.client

import openpyxl
import requests
from bs4 import BeautifulSoup
import time
import os

from datetime import datetime


excel_file_path = "" // #Label file path
workbook = openpyxl.load_workbook(excel_file_path)
sheet = workbook.active


chrome_options = Options()
chrome_options.add_argument("--headless") 
chrome_options.add_argument("--disable-gpu") 


try:
    s = Service(ChromeDriverManager().install())
    browser = webdriver.Chrome(service=s)
except Exception as e:
    print(e)
    exit()


url = "" #website URL
browser.get(url)


def click(a):
    search_box = browser.find_element(By.CSS_SELECTOR, a)
    search_box.click()


click("#header > div > p > a")
time.sleep(2)

company_input = browser.find_element(By.CSS_SELECTOR, "#CompanyId")
company_input.click()
company_input.send_keys("") #ID


login_name_input = browser.find_element(By.CSS_SELECTOR, "#LoginName")
login_name_input.click()
login_name_input.send_keys("") #Password

pass_input = browser.find_element(By.CSS_SELECTOR, "#Passwd")
pass_input.click()
pass_input.send_keys("Hsn07072!")

click("#aSignIn")
time.sleep(1)

History = "
browser.get(History)


initial_id = 1

while True:
    invoice = input("Put invoice (or 'N' to exit): ")
    if invoice.upper() == 'N':
        break

    invoice_element = browser.find_element(By.CSS_SELECTOR, "#UniqueId")
    invoice_element.click()
    invoice_element.clear()
    invoice_element.send_keys(invoice)

    click("#Refresh") 

    time.sleep(1)

    while True:
        td_id = f"data_{initial_id}_0_0"
        
        try:
            td_element = browser.find_element(By.ID, td_id)
            

            date_element = td_element.find_element(By.TAG_NAME, "span")
            date = date_element.text

            anchor_selector = 'a[href^="javascript:showSsPartInfo"]'
            part_element = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, anchor_selector)))
            part = part_element.text

            print(invoice)
            print(date)
            print(part)
            
            initial_id += 2  

            sheet['B2'] = invoice
            sheet['B3'] = part
            sheet['F5'] = date
            workbook.save(excel_file_path)

            os.startfile(excel_file_path, "print")

            
        except:
             print(f"No element found with the current td_id: {td_id}")

            break
            
        
