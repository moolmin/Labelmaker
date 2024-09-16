from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import openpyxl
import time
import os

# Path to the Excel file
excel_file_path = ""  # Label file path
workbook = openpyxl.load_workbook(excel_file_path)
sheet = workbook.active

# Set Chrome options for headless mode
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")

# Initialize the Chrome driver
try:
    s = Service(ChromeDriverManager().install())
    browser = webdriver.Chrome(service=s)
except Exception as e:
    print(e)
    exit()

# URL of the website to automate
url = ""  # website URL
browser.get(url)

# Function to click on a web element specified by CSS selector
def click(a):
    search_box = browser.find_element(By.CSS_SELECTOR, a)
    search_box.click()

# Click on the login button
click("#header > div > p > a")
time.sleep(2)

# Enter company ID
company_input = browser.find_element(By.CSS_SELECTOR, "#CompanyId")
company_input.click()
company_input.send_keys("")  # ID

# Enter login name
login_name_input = browser.find_element(By.CSS_SELECTOR, "#LoginName")
login_name_input.click()
login_name_input.send_keys("")  # Password

# Enter password
pass_input = browser.find_element(By.CSS_SELECTOR, "#Passwd")
pass_input.click()
pass_input.send_keys("password")

# Click on the sign-in button
click("#aSignIn")
time.sleep(1)

# Navigate to the history page
History = ""
browser.get(History)

# Initial ID for the table data
initial_id = 1

# Loop to continuously accept invoice numbers until 'N' is entered
while True:
    invoice = input("Put invoice (or 'N' to exit): ")
    if invoice.upper() == 'N':
        break

    # Find and enter the invoice number
    invoice_element = browser.find_element(By.CSS_SELECTOR, "#UniqueId")
    invoice_element.click()
    invoice_element.clear()
    invoice_element.send_keys(invoice)

    # Click the refresh button
    click("#Refresh")
    time.sleep(1)

    # Loop to find table data with increasing IDs
    while True:
        td_id = f"data_{initial_id}_0_0"
        
        try:
            td_element = browser.find_element(By.ID, td_id)
            
            # Extract date from the table cell
            date_element = td_element.find_element(By.TAG_NAME, "span")
            date = date_element.text

            # Extract part information using a specific anchor selector
            anchor_selector = 'a[href^="javascript:showSsPartInfo"]'
            part_element = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, anchor_selector)))
            part = part_element.text

            # Print the invoice, date, and part information
            print(invoice)
            print(date)
            print(part)
            
            initial_id += 2  # Increment ID for next table cell

            # Write data to the Excel sheet and save
            sheet['B2'] = invoice
            sheet['B3'] = part
            sheet['F5'] = date
            workbook.save(excel_file_path)

            # Open the Excel file for printing
            os.startfile(excel_file_path, "print")

        except:
            print(f"No element found with the current td_id: {td_id}")
            break
