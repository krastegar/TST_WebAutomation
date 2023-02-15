import time 

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By


class SetUp:
    def __init__(
        self, 
        username, 
        paswrd, 
        FromDate, 
        ToDate,
        url='https://test-sdcounty.atlasph.com/TSTWebCMR/pages/login/login.aspx'
        ):
        self.url = url 
        self.username = username
        self.paswrd = paswrd
        self.fromDate = FromDate
        self.toDate = ToDate
    
    def install(self): 
        '''
        If webdriver is not installed already this function will install it 
        via python commands 
        '''
        service = ChromeService(executable_path=ChromeDriverManager().install())
        return


    def login(self): 

        '''
        Function is meant to go to specified url for webcmr, i.e)
        TSTWebCMR
        TRNWebCMR
        WebCMR (Production)
        '''
        driver = webdriver.Chrome()
        website = driver.get(self.url)
        # Find the username and password elements and enter login credentials
        time.sleep(1)
        username = driver.find_element(By.ID, value="txtUsername")
        username.send_keys(self.username)
        password = driver.find_element(By.ID, value="txtPassword")
        password.send_keys(self.paswrd)
        time.sleep(.5)
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
