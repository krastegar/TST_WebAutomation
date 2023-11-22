import time
import os 

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By


class SetUp:
    def __init__(
        self, 
        username, 
        paswrd, 
        FromDate, 
        ToDate,
        url = 'https://test-sdcounty.atlasph.com/TSTWebCMR/pages/login/login.aspx'
        ):
        self.url = url 
        self.username = username
        self.paswrd = paswrd
        self.fromDate = FromDate
        self.toDate = ToDate


    def login(self): 

        '''
        Function is meant to go to specified url for webcmr, i.e)
        TSTWebCMR
        TRNWebCMR
        WebCMR (Production)
        '''
        # create chrome webdriver object with the above options
        service = ChromeService(executable_path="chromedriver.exe")
        driver = webdriver.Chrome(service=service)

        # go to TST website
        driver.get(self.url)

        # Find the username and password elements and enter login credentials
        # time.sleep(30)
        wait = WebDriverWait(driver, 30)
        username = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtUsername"]')))
        username.send_keys(self.username)

        password = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtPassword"]')))
        password.send_keys(self.paswrd)
        password.send_keys(Keys.RETURN)
        
        return driver

    def date_range(self,driver): 
        # inputting date range 
        from_dt, to_dt =  'txtFromDate', 'txtToDate'
        fromDate = driver.find_element(By.ID, from_dt)
        fromDate.send_keys(self.fromDate)
        toDate = driver.find_element(By.ID, to_dt)
        toDate.send_keys(self.toDate)
        return
