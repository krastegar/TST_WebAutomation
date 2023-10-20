import os
import time
import glob
import pandas as pd
# import numpy as np 

from SetUp import SetUp
# from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException


class IMM(SetUp): 
    
    '''
    For now I am going to do double inheritance, because I am making the assumption
    that the date range for staging area will be the same as the date range for 
    incoming message monitor 
    '''
    
    def __init__(self, LabName : str, *args, **kwargs): 
        super().__init__(*args, **kwargs)
        self.lab = LabName
        

    def imm_menu(self): 
        '''
        This method is used for logging into TSTWebCMR and going to 
        Incoming Message Monitor page 
        '''
        # login and go to home page
        driver = self.login()
        # Navigating the dropdown menu to go into IMM
        self.nav2IMM(driver)
        return driver 
    
    def go_home(self, driver):
        '''
        Method is meant to navigate back home, using home icon button
        to get there
        '''
        # Clicking home button 
        home_btn = self.multiFind(
            driver = driver, 
            element_id='FragTop1_lbtnHome',
            xpath= '/html/body/form/table[2]/tbody/tr/td[1]/div/a'
            )
        home_btn.click()

        # Clicking Search button (this is the investigators equivalent of my home button)
        investigators_search = self.multiFind(
            driver=driver,
            element_id='FragTop1_mnuMain-menuItem002',
            xpath='/html/body/form/table[2]/tbody/tr/td[2]/table[36]/tbody/tr/td[1]'
        )
        investigators_search.click()
        return
    
    def nav2IMM(self, driver):
        '''
        Only works if you are at the home page and want to navigate to IMM menu
        '''
        # Navigating back to home page
        self.go_home(driver)

        # Looking at Administrator dropdown menu
        wait = WebDriverWait(driver, 8)
        #dropdown_menu = wait.until(EC.presence_of_element_located((By.ID, "FragTop1_mnuMain-menuItem017")))
        dropdown_menu = self.multiFind(
            driver=driver,
            element_id= "FragTop1_mnuMain-menuItem017",
            xpath='/html/body/form/table[2]/tbody/tr/td[2]/table[36]/tbody/tr/td[6]'
        )
        dropdown_menu.click()

        # Hopefully clicking on incoming message monitor options
        # Wait for the second level dropdown to be present and then click on it
        #second_menu = "FragTop1_mnuMain-menuItem017-subMenu-menuItem009"
        second_level_dropdown = self.multiFind(
            driver, 
            element_id= "FragTop1_mnuMain-menuItem017-subMenu-menuItem009",
            xpath='/html/body/form/table[2]/tbody/tr/td[2]/table[16]/tbody/tr[9]/td'
        )
        second_level_dropdown.click()

        # Wait for the desired option to be present and then click on it
        imm = 'FragTop1_mnuMain-menuItem017-subMenu-menuItem009-subMenu-menuItem003'
        desired_option = self.multiFind(
            driver=driver,
            element_id=imm,
            xpath='/html/body/form/table[2]/tbody/tr/td[2]/table[5]/tbody/tr[3]/td'
        )
        desired_option.click()
        return
    
    def imm_search(self): 
        '''
        This method is for conducting the search of specific dates and a specific lab that 
        we want to do validations for. (This is the reason for double inheritance)
        '''
        # Going to IMM page
        driver = self.imm_menu()

        # inputting date ranges 
        input_dates = self.date_range(driver)
        
        # Choosing lab of interest
        imm_lab_menu = 'ddlLaboratory'
        lab_dropdown = self.multiFind(
            driver=driver,
            element_id=imm_lab_menu,
            xpath='/html/body/form/div[3]/div/div/table[3]/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/div[1]/select'
        )
        lab_dropdown.click()
        
        # Create a Select object and select Lab of Interest 
        select = Select(lab_dropdown)
        select.select_by_visible_text(self.lab)
        
        # Find search button and click 
        search_id = 'ibtnSearch'
        search =  self.multiFind(
            driver=driver,
            element_id=search_id,
            xpath='/html/body/form/div[3]/div/div/table[3]/tbody/tr[2]/td/table/tbody/tr[4]/td/div/input[1]'
        )
        search.click()

        return driver

    def imm_export(self): 
        '''
        Method is used after completing the search for specific lab and exporting
        all the results as an excel workbook and looking up those results in TST system.
        **Takes you back to the home page after excel workbook is downloaded
        '''
        driver = self.imm_search()
        
        # clicking export button
        csv_id = 'btnExport_btnExport'
        wait = WebDriverWait(driver, 10)
        csv_sheet = wait.until(EC.presence_of_element_located((By.ID, csv_id)))
        csv_sheet.click()

        # making selections in pop-up menu
        iframe = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/iframe")))
        driver.switch_to.frame(iframe)

        # choosing select all for columns
        select_all = driver.find_element(By.ID, 'clstAll_0' )
        select_all.click()

        # choosing microsoft excel format
        csv_option = driver.find_element(by=By.ID, value='optFormat_2')
        csv_option.click()

        # Clicking Export
        export_id = 'btnExport'
        export_btn = wait.until(EC.presence_of_element_located((By.ID, export_id)))
        export_btn.click()

        # switching back to parent frame / default content: 
        driver.switch_to.default_content()

        # Making relative path to Downloads folder
        download_folder = self.download_folder()
        initial_files = set(os.listdir(download_folder))
        
        # Wait for the download to complete
        while True: 
            current_files = set(os.listdir(download_folder))
            if len(current_files) <= len(initial_files): 
                try:
                    # going to wait 6 seconds if nothing shows up 
                    # will wait 6 more seconds
                    time.sleep(6)
                except: 
                    pass 
            elif len(current_files) > len(initial_files): 
                break
            else: 
                break 
        # driver.quit()
        self.go_home(driver)
        return driver

    def download_folder(self):
        '''
        Method to get the relative path of a users Downloads folder
        '''
        home_directory = os.path.expanduser("~")
        download_folder = os.path.join(home_directory, "Downloads")
        return download_folder

    def export_df(self): 
        
        download_folder = self.download_folder()
        
        # Reading in the most recently downloaded excel file from downloads folder
        list_of_files = glob.glob(f'{download_folder}/*.xlsx')
        latest_file = max(list_of_files, key=os.path.getctime)
        export_df = pd.read_excel(latest_file)

        # Filtering results based on unique resulted tests and test type
        # Remember this is only for imported reports ONLY not all 
        imported_status_df = export_df[export_df['DILR_ImportStatus']=='Imported']
        filtered_import = imported_status_df[['DILR_ResultedTest', 
                                                'DILR_IncidentID', 
                                                'DILR_AccessionNumber',
                                                'DILR_ResultValue']]
        
        # only looking at last unique result value for each test
        # This will make sure that the resultValue and resultTest are unique combos
        filtered_import.drop_duplicates(subset=['DILR_ResultedTest','DILR_ResultValue'],
                                                        keep='last',
                                                        inplace=True)

        # Need to classify each result so that I can filter the numeric results and just pick 
        # one accession ID and disease ID for each test that has numeric
        # Will do the same for mixed as well 
        filtered_import['Classification'] = filtered_import['DILR_ResultValue'
                                                        ].apply(self.classify_column)
        
        # Final filter of df based on classification column and ResultedTest
        # Although we kept all of the unique combinations of DILR_ResultValue and ResultedTest after classification
        # We might have multiple numeric values and we would not want to look at every single one so instead we
        # only look at one example of this classification for each ResultedTest and that is also why we split the 
        # data into two dataframes and recombine them (because we only want one example of ERROR and Numeric)
        
        num_error = filtered_import[filtered_import['Classification'].isin(['ERROR', 'Numeric'])]
        num_error.drop_duplicates(subset=['DILR_ResultedTest','Classification'],
                                                        keep='last',
                                                        inplace=True)
        categorical_df = filtered_import[filtered_import['Classification'].isin(['Categorical'])]
        
        # Combine dataframe 
        final_df = pd.concat([categorical_df, num_error], ignore_index=True)
        final_df.to_excel('ELR_Validation_Search_Summary.xlsx')

        return final_df

    def classify_column(self,value):
        '''
        Method is meant to be used to classify ResultValue column of dataframe
        and split up the weird results away from the normal ones 
        '''
        # Check if value contains only letters
        if value.isalpha():
            # If value contains only letters, return 'Categorical'
            return 'Categorical'
        else:
            # If value does not contain only letters, try to convert value to float
            try:
                float(value)
                # If successful, return 'Numeric'
                return 'Numeric'
            except:
                # If not successful, check if value contains a mixture of letters and digits
                if any(char.isalpha() for char in value) and any(char.isdigit() for char in value):
                    # If value contains a mixture of letters and digits, return 'Mixed'
                    return 'ERROR'
                else:
                    # Return 'Categorical' for other cases
                    return 'Categorical'
    def multiFind(self, driver, element_id, xpath=None, field_name=None):
         
        '''
        Function is meant to locate regions on html web page, 
        using the elements ID as a identifier. If element is not found 
        by ID, function will attempt to find it by XPath.
        '''
        
        wait = WebDriverWait(driver, 1)
        try: 
            element_btn = wait.until(EC.presence_of_element_located((By.ID, element_id)))
        except TimeoutException:
            #time.sleep(3)
            if xpath:
                try: 
                    element_btn = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                except: 
                    pass
            if field_name:
                try: 
                    element_btn = wait.until(EC.presence_of_element_located((By.XPATH, field_name)))
                except: 
                    pass  
            else: print(f'Cannot locate desired field or tab')
        except ElementClickInterceptedException: 
            self.alert_handling(driver)

        return element_btn  