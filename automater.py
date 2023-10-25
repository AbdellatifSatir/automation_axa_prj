from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import datetime
import time

    #   THIS SCRIPT IS FOR AUTOMATE THE DAILY DATA EXTRACTION   #

URL = '' #private
CHROME_PATH = r'.\chromedriver_win32\chromedriver.exe' # chrome driver path
CHROME_SERVICE = ChromeService(executable_path=CHROME_PATH)
CHROME_OPTIONS = ChromeOptions()
CHROME_OPTIONS.add_argument('--headless=new') # background mode
CHROME_DRIVER = webdriver.Chrome(service=CHROME_SERVICE, options=CHROME_OPTIONS)
CHROME_DRIVER.get(URL)
print('\tWaiting for the page be visible...')

# Wait for an element to be visible before proceeding
WAIT = WebDriverWait(CHROME_DRIVER, 10)
CURRENT_DATE = datetime.date.today()
DAY_NAME = CURRENT_DATE.strftime("%A")
CURRENT_DATE_FORMATED = CURRENT_DATE.strftime('%d/%m/%Y')
N_DAYS = (1 if DAY_NAME != 'Monday' else 3) # weekend
CURRENT_DATE_MINUS_DAYS = (CURRENT_DATE - datetime.timedelta(days=N_DAYS))
CURRENT_DATE_MINUS_DAYS_FORMATED = CURRENT_DATE_MINUS_DAYS.strftime('%d/%m/%Y')

    # SELCET OPTION FIELD #
try:
    SELECT_ELEMENT = WAIT.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ReportList"]')))
    SELECT = Select(SELECT_ELEMENT)
    SELECT.select_by_visible_text('GEDS')
except:
    print('There is no option selected')

    # DATE FIELD #
try:
    StartDateComment = WAIT.until(EC.visibility_of_element_located((By.ID, 'StartDateBox1')))
    StartDateComment.send_keys(CURRENT_DATE_MINUS_DAYS_FORMATED)
    EnDateComment = WAIT.until(EC.element_to_be_clickable((By.ID, 'EndDateBox2')))
    EnDateComment.send_keys(CURRENT_DATE_FORMATED)
except:
    EnDateComment = WAIT.until(EC.element_to_be_clickable((By.ID, 'EndDateBox2')))
    EnDateComment.send_keys(CURRENT_DATE_FORMATED)
print('\tWaiting for date entry...')

    # SUBMIT FIELD #
try:
    print('\tStarting download the data...')
    SUBMIT = WAIT.until(EC.element_to_be_clickable((By.ID, 'Execute')))
    CHROME_DRIVER.execute_script("arguments[0].scrollIntoView();", SUBMIT)
    CHROME_DRIVER.execute_script("arguments[0].click();", SUBMIT)
except:
    print('The form does not submited')
print('\tFinish downloading check the download folder.')

time.sleep(2)
CHROME_DRIVER.quit()
print('\tPress any key to exit.')