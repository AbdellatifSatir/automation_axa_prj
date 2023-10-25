from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import datetime
import time

    # THIS SCRIPT IS FOR EXTRACT 2023 DATA #

URL = 'http://pragdgencti01.applications.services.axa-tech.intraxa/reporting-wfm/reporting_wfm.aspx'
CHROME_PATH = r'.\chromedriver_win32\chromedriver.exe'
CHROME_SERVICE = ChromeService(executable_path=CHROME_PATH)
CHROME_OPTIONS = ChromeOptions()
CHROME_OPTIONS.add_argument('--headless=new')
CHROME_DRIVER = webdriver.Chrome(service=CHROME_SERVICE, options=CHROME_OPTIONS)

START_DATE = datetime.datetime(2023, 1, 1)
END_DATE = datetime.datetime(2023, 7, 1)

END_DATE_FORMATED = END_DATE.strftime('%d/%m/%Y')
START_DATE_FORMATED = START_DATE.strftime('%d/%m/%Y')
DELTA = datetime.timedelta(days=6)
D = datetime.timedelta(days=0)

while (START_DATE <= END_DATE):

    if (START_DATE + DELTA) >=  END_DATE:
        START_DATE = END_DATE - DELTA

    print('\n', START_DATE.strftime('%d/%m/%Y'), 'to', (START_DATE + DELTA).strftime('%d/%m/%Y'))
    CHROME_DRIVER.get(URL)
    print('\tWaiting for the page be visible...')
    WAIT = WebDriverWait(CHROME_DRIVER, 10)

        # SELECT FIELD #
    try:
        SELECT_ELEMENT = WAIT.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ReportList"]')))
        SELECT = Select(SELECT_ELEMENT)
        SELECT.select_by_visible_text('GEDS')
    except:
        print('There is no option selected')

        # DATE FIELD #
    try:
        StartDateComment = WAIT.until(EC.visibility_of_element_located((By.ID, 'StartDateBox1')))
        StartDateComment.send_keys((START_DATE).strftime('%d/%m/%Y'))
        EnDateComment = WAIT.until(EC.element_to_be_clickable((By.ID, 'EndDateBox2')))
        EnDateComment.send_keys((START_DATE + DELTA).strftime('%d/%m/%Y'))
    except:
        EnDateComment = WAIT.until(EC.element_to_be_clickable((By.ID, 'EndDateBox2')))
        EnDateComment.send_keys((START_DATE + DELTA).strftime('%d/%m/%Y'))
    print('\tWaiting for date entry...')

        # SUBMIT FIELD #
    try:
        time.sleep(1)
        SUBMIT = WAIT.until(EC.element_to_be_clickable((By.ID, 'Execute')))
        CHROME_DRIVER.execute_script("arguments[0].scrollIntoView();", SUBMIT)
        CHROME_DRIVER.execute_script("arguments[0].click();", SUBMIT)
        print('\tStarting download the data...')
        time.sleep(15)
    except:
        print('The form does not submited')

    time.sleep(2)
    START_DATE = START_DATE + DELTA + datetime.timedelta(days=1)
    time.sleep(2)
    print('\tNext Waiting for new date entry...')

print('\nFinish downloading check the download folder.')
# CHROME_DRIVER.quit()