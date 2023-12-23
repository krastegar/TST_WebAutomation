"""
The SetUp class is responsible for setting up the necessary parameters and performing 
actions related to establishing webdriver sessions, logging in, and manipulating date ranges 
on a website.

Algorithm:
1. The __init__ method initializes the SetUp object with the provided username, password, start date, end date, and an optional URL parameter.
2. The login method logs into a website using the provided username and password, and returns a WebDriver object.
3. The date_range method inputs the date range in the provided WebDriver object.
"""

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
        """
            Login to the website using the provided username and password.
            
            :return: The webdriver object after successful login.
            :rtype: WebDriver
        """

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
        """
        Input the date range in the provided driver object.
        
        :param driver: The driver object used to interact with the web page.
        :type driver: WebDriver
        
        :return: None
        """
        
        fromDate = driver.find_element(By.ID, 'txtFromDate')
        fromDate.send_keys(self.fromDate)
        toDate = driver.find_element(By.ID, 'txtToDate')
        toDate.send_keys(self.toDate)
        return
