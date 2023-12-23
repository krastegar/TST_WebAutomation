import sys
import io 

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException
import time
import pandas as pd 

def disease_search():
    filename = r"C:\Users\krastega\Downloads\777291b2-4d0b-4fc0-836b-47400d10e67e.xlsx"
    patient_data = pd.read_excel(filename)
    filtered_df = patient_data[(patient_data['DILR_ResultTest']=='Hepatitis B core, IgM - Final') & (patient_data['DILR_ImportStatus']=='Imported')]
    DI_incident = list(filtered_df['DILR_IncidentID'])
    # query = int(DI_incident[0])
    return DI_incident

def main():
    
    # Create a new instance of the chrome driver
    # service = ChromeService(executable_path=ChromeDriverManager().install())
    driver = webdriver.Chrome()

    driver.get("https://test-sdcounty.atlasph.com/TSTWebCMR/pages/login/login.aspx")
    # Find the username and password elements and enter login credentials
    time.sleep(0.5)
    username = driver.find_element(By.ID, value="txtUsername")
    username.send_keys("krastegar")
    password = driver.find_element(By.ID, value="txtPassword")
    password.send_keys("Hamid&Mahasty1")
    time.sleep(0.5)
    password.send_keys(Keys.RETURN)
   
   # Navigating the dropdown menu to go into IMM
    time.sleep(0.5)
    wait = WebDriverWait(driver, 8)
    dropdown_menu = wait.until(EC.presence_of_element_located((By.ID, "FragTop1_mnuMain-menuItem017")))
    dropdown_menu.click()

    # Hopefully clicking on incoming message monitor options
    # Wait for the second level dropdown to be present and then click on it
    time.sleep(0.5)
    second_menu = "FragTop1_mnuMain-menuItem017-subMenu-menuItem009"
    second_level_dropdown = wait.until(EC.presence_of_element_located((By.ID, second_menu)))
    second_level_dropdown.click()

    # Wait for the desired option to be present and then click on it
    imm = 'FragTop1_mnuMain-menuItem017-subMenu-menuItem009-subMenu-menuItem003'
    desired_option = wait.until(EC.presence_of_element_located((By.ID, imm)))
    desired_option.click()

    # input date of incoming messages 
    time.sleep(0.5)
    DOI = '01/18/2023'
    imm_dateID = 'txtFromDate'
    imm_date = driver.find_element(by=By.ID, value=imm_dateID)
    imm_date.send_keys(DOI)
    
    time.sleep(0.5)
    imm_ToDateID = 'txtToDate'
    imm_ToDate = driver.find_element(By.ID, imm_ToDateID)
    imm_ToDate.send_keys(DOI)
    time.sleep(0.5)

    # Choosing Lab of Interest
    lab_drpdown_menu = 'ddlLaboratory'
    lab_dropdown = wait.until(EC.presence_of_element_located((By.ID, lab_drpdown_menu)))
    lab_dropdown.click()
    time.sleep(0.5)
    
    # Create a Select object
    select = Select(lab_dropdown)

    # Select the option with the desired value
    logan_heights = '980139'
    select.select_by_value(logan_heights)

    # conduct search 
    time.sleep(0.5)
    lab_search_id= 'ibtnSearch'
    search_button = driver.find_element(by=By.ID, value=lab_search_id)
    search_button.click()

    # Clicking on export button which produces a pop up table
    time.sleep(0.5)
    csv_id = 'btnExport_btnExport'
    csv_sheet = wait.until(EC.presence_of_element_located((By.ID, csv_id)))
    csv_sheet.click()

    time.sleep(0.5)
    # Trying to click on an element inside pop up menu (this produces the errors)
    # Switch to iframe that contains pop-up table
    iframe = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/iframe")))
    driver.switch_to.frame(iframe)
    time.sleep(0.5)
    select_all = driver.find_element(By.ID, 'clstAll_0' )
    select_all.click()
    time.sleep(0.5)
    csv_option = driver.find_element(by=By.ID, value='optFormat_2')
    csv_option.click()
    time.sleep(0.5)
    export_btn = driver.find_element(By.ID, 'btnExport')
    export_btn.click()
    driver.switch_to.default_content() # getting out of iframe and back to main html

    # go back to homepage
    time.sleep(1)
    home_btn = driver.find_element(By.ID, 'FragTop1_lbtnHome')
    home_btn.click()

    # click on disease incident
    time.sleep(1)
    DI_tab = driver.find_element(By.ID, 'ibtnTabDisInc')
    DI_tab.click()

    # Getting DI ID to produce pdf report 
    time.sleep(1)
    incindents = disease_search() # array of queries
    print(incindents)
    for i in incindents:
        '''
        After conducting the first search for the disease incident, the search bar
        becomes stale. This is why we implement the try-except block to catch 
        StaleElementReferenceException 
        '''
        try:
            DI_search = driver.find_element(By.ID, 'txtFindDisInc')
            DI_search.clear() 
            DI_search.send_keys(int(i))
            time.sleep(2)
            DI_search.send_keys(Keys.RETURN)
            time.sleep(2)
            DI_search.clear()
        except StaleElementReferenceException as e:
            print("An error occurred:", e)
            # Refresh the reference to the search bar element
            DI_search = driver.find_element(By.ID, 'txtFindDisInc')
            DI_search.clear() 






    # keep browser open as long as I want 
    input("Press enter to close Browser")
    driver.quit()

    return

if __name__=="__main__":
    sys.exit(main())