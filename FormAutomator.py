import time
import pandas as pd
import logging
import subprocess
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Logging setup
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s [%(levelname)s] %(message)s")

EXCEL_FILE = r"C:\Users\suraj\Downloads\IQOO_CGO_Data_Formatted split(SK).xlsx"
CHROMEDRIVER_PATH = r"C:\Users\suraj\Downloads\chromedriver-win64\chromedriver.exe"
STATUS_LOG_FILE = r"C:\\Users\\suraj\\Downloads\\submission_log.txt"

# Clear previous log
open(STATUS_LOG_FILE, "w", encoding="utf-8").close()

# Load Excel
df = pd.read_excel(EXCEL_FILE)
logging.info(f"Loaded {len(df)} rows from Excel file.")

# Setup driver
service = Service(CHROMEDRIVER_PATH)
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 15)

for index, row in df.iterrows():
    try:
        logging.info(f"Processing row {index + 1} - {row['First Name']} {row['Last Name']}")
        driver.get("https://forms.zohopublic.in/the23watts1/form/iQOOsChiefGamingOfficer/formperma/9P8OLeRo5RHEcvU0BUzaDF_w2Rsd0KRylCLde1NHv78")

        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="Name-li"]//input'))).send_keys(str(row['First Name']))
        driver.find_element(By.XPATH, '//*[@id="Name-li"]//div[2]/span/input').send_keys(str(row['Last Name']))
        time.sleep(1)

        driver.find_element(By.XPATH, '//*[@id="Email-li"]//input').send_keys(str(row['Email']))
        time.sleep(1)

        driver.find_element(By.XPATH, '//*[@id="Date-date"]').send_keys(str(row['Date of Birth']))
        time.sleep(1)

        # City Dropdown
        driver.find_element(By.XPATH, '//*[@id="Dropdown-li"]//span[contains(@class,"select2-selection")]').click()
        city_input = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@class="select2-search__field"]')))
        city_input.send_keys(str(row['City']))
        time.sleep(1)
        options_list = driver.find_elements(By.XPATH, '//li[contains(@class,"select2-results__option") and not(contains(@class,"loading"))]')
        if options_list:
            matched = False
            for opt in options_list:
                if opt.text.strip().lower() == row['City'].strip().lower():
                    opt.click()
                    matched = True
                    break
            if not matched:
                for opt in options_list:
                    if opt.text.strip().lower() == "other":
                        opt.click()
                        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="Dropdown-li"]//input[@type="text"]'))).send_keys(str(row['City']))
                        break
        time.sleep(1)

        # State Dropdown
        driver.find_element(By.XPATH, '//*[@id="Dropdown1-li"]//span[contains(@class,"select2-selection")]').click()
        state_input = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@class="select2-search__field"]')))
        state_input.send_keys(str(row['State']))
        time.sleep(1)
        options_list = driver.find_elements(By.XPATH, '//li[contains(@class,"select2-results__option") and not(contains(@class,"loading"))]')
        if options_list:
            matched = False
            for opt in options_list:
                if opt.text.strip().lower() == row['State'].strip().lower():
                    opt.click()
                    matched = True
                    break
            if not matched:
                for opt in options_list:
                    if opt.text.strip().lower() == "other":
                        opt.click()
                        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="Dropdown1-li"]//input[@type="text"]'))).send_keys(str(row['State']))
                        break
        time.sleep(1)

        driver.find_element(By.XPATH, '//*[@id="Website-arialabel"]').send_keys(str(row['Instagram Handle Link']))
        time.sleep(1)

        driver.find_element(By.XPATH, '//*[@id="PhoneNumber"]').send_keys(str(row['Phone']))
        time.sleep(1)

        # How did you hear Dropdown
        driver.find_element(By.XPATH, '//*[@id="Dropdown2-li"]//span[contains(@class,"select2-selection")]').click()
        source_input = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@class="select2-search__field"]')))
        source_input.send_keys("Other")
        option = wait.until(EC.presence_of_element_located((By.XPATH, '//li[contains(@class,"select2-results__option") and not(contains(@class,"loading"))]')))
        option.click()
        time.sleep(1)

        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="Dropdown2-li"]//input[@type="text"]'))).send_keys("RS")
        time.sleep(1)

        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="formBodyDiv"]//button'))).click()
        logging.info(f"Row {index + 1} submitted successfully.")
        time.sleep(2)

        with open(STATUS_LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{index + 1}, {row['First Name']} {row['Last Name']}, SUCCESS\n")

    except Exception as e:
        logging.error(f"Error on row {index + 1}: {e}")
        with open(STATUS_LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{index + 1}, {row['First Name']} {row['Last Name']}, ERROR: {str(e)}\n")

logging.info("All rows processed. Quitting browser.")
driver.quit()

# Submission summary
success_count = 0
error_count = 0

with open(STATUS_LOG_FILE, "r", encoding="utf-8") as f:
    for line in f:
        if "SUCCESS" in line:
            success_count += 1
        elif "ERROR" in line:
            error_count += 1

summary = f"\nSubmission Summary:\nSUCCESS: {success_count}\nERRORS: {error_count}"
print(summary)
logging.info(summary)

# Open the log file
subprocess.Popen(["notepad.exe", STATUS_LOG_FILE])