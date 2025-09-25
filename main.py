import time
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import glob
import pandas as pd


URL = "https://proapp.techway.online/index.aspx"
USERNAME = "*********************"
PASSWORD = "***********************"
downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")

yesterday = datetime.now() - timedelta(days=1)
date_full = yesterday.strftime("%d-%b-%Y")   # e.g., 21-Sep-2025
date_short = yesterday.strftime("%d-%b-%y")  # e.g., 21-Sep-25

print("Using dates:", date_full, "and", date_short)


options = Options()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)

try:
    driver.get(URL)
    wait = WebDriverWait(driver, 15)
    username_el = wait.until(EC.presence_of_element_located((By.ID, "txtUserName")))
    password_el = wait.until(EC.presence_of_element_located((By.ID, "txtPassword")))
    username_el.clear()
    username_el.send_keys(USERNAME)
    password_el.clear()
    password_el.send_keys(PASSWORD)
    login_btn = wait.until(EC.element_to_be_clickable((By.ID, "ImgSubmit")))
    login_btn.click()
    periodic_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Periodic Production")))
    periodic_link.click()
    from_date_el = wait.until(EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_txtfromdate")))
    to_date_el = wait.until(EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_txttodate")))
    from_date_el.clear()
    from_date_el.send_keys(date_short)   # 21-Sep-25
    to_date_el.clear()
    to_date_el.send_keys(date_full)      # 21-Sep-2025
    print("✅ Dates entered:", date_short, "to", date_full)
    detail_btn = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_lnkDetails")))
    detail_btn.click()
    print("✅ Clicked 'Detail Prod. Report'")
    time.sleep(5)
    export_btn = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnExport")))

    export_btn.click()
    print("✅ Export button clicked, waiting for download...")
    time.sleep(10)
    list_of_files = glob.glob(os.path.join(downloads_folder, "*.xls"))
    if list_of_files:
        file_path = max(list_of_files, key=os.path.getctime)  
    time.sleep(5)
    print(f"✅ File downloaded: {file_path}")

    tables = pd.read_html(file_path, flavor='html5lib')
    df = tables[0]
    df_fg = df[df['Item Category'] == 'FG']
    df_fg.to_excel("haldiaeol.xlsx", index=False)

    df_bb = df[df['Item Category'].str.contains('BB', na=False)]
    df_bb.to_excel("haldiabulkbag.xlsx", index=False)

    print("✅ Two files created: haldiaeol.xlsx and haldiabulkbag.xlsx")


finally:
    driver.quit()
