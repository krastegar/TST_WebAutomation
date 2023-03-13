import time
import os
import pandas as pd
import re
import logging

from IMM import IMM
from SetUp import SetUp
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

class DI_Search(IMM, SetUp): 

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
    
    def disease_search(self): 
        # Going into IMM Menu and Downloading excel sheet 
        driver = self.imm_export()

        # Getting Disease Queries and Hl7 accession # from excel sheet
        df = self.export_df()
        di_num_list = list(df['DILR_IncidentID'])

        # click home button 
        self.go_home(driver)
        webCMR_values, webCMR_indicies = None, None # bounding these variables
        
        # Scraping values from Demographic and Lab tabs off of TST website
        for i in range(len(di_num_list)):
            webCMR_values, webCMR_indicies = self.webTST_scrape(driver, di_num_list, i)
        
        return driver, df, webCMR_values, webCMR_indicies

    def webTST_scrape(self, driver, di_num_list, i):
        di_num = int(di_num_list[i])
            
        # navigating to disease incidents tab
        logging.info('Clicking Disease Incident tab')
        DI_tab = driver.find_element(By.ID, 'ibtnTabDisInc')
        DI_tab.click()
            
        # Searching for specific disease incident 
        logging.info("Inputting DiseaseIncidentID and clicking on Results ")
        di_search_box = 'txtFindDisInc'
        DI_search = driver.find_element(By.ID, di_search_box)
        DI_search.send_keys(str(di_num))
        DI_search.send_keys(Keys.RETURN)
        di_link = driver.find_element(By.LINK_TEXT, str(di_num))
        di_link.click()
            
            # to by pass alert about locked case (just need to read stuff)
        try: 
            alert = Alert(driver)
            alert.accept()
        except:
            pass

            # grabbing information from boxes of interest 
            
            # ----- Demographics tab -----------
            logging.info("Extracting Data from Demographic Tab")
            # Name fields
        last_name = self.extract_info(
            driver = driver, 
            element_id='txtLastName',
            #'/html/body/form/div[2]/div/div/table/tbody/tr/td/table[2]/tbody/tr[3]/td/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div/div[2]/div[1]/input',
            field_name='Last Name'
            )
        first_name = self.extract_info(
            driver = driver, 
            element_id='txtFirstName', 
            #'/html/body/form/div[2]/div/div/table/tbody/tr/td/table[2]/tbody/tr[3]/td/table/tbody/tr[2]/td[1]/table/tbody/tr[1]/td/div/div[2]/div[2]/input',
            field_name='First Name'
            )
        # print(f"FirstName  FIELD: {first_name} \n Data Type: {type(first_name)}")
        # print(f"LastName  FIELD: {last_name} \n Data Type: {type(last_name)}")
        name = first_name + ' ' + last_name

        #print(f'Demographic NAME: {name}')

            # DOB
        dob = self.extract_info(
            driver=driver, 
            element_id='txtDOB',
            field_name='DOB'
            )

            # Reported Race
        race = self.extract_info(
            driver=driver, 
            element_id='txtReportedRace',
            field_name='Race'
            )

            # Ethnicity
        ethnicity = self.dropDown_extract(driver, 'cboEthnicity')

            # Address
        street = self.extract_info(
            driver = driver, 
            element_id='txtAddress',
            field_name='Address'
            )
        unit = self.extract_info(
            driver = driver, 
            element_id='txtApartment',
            field_name='Apartment'
            )
        city = self.extract_info(
            driver=driver, 
            element_id='txtCity',
            field_name='City'
            )
        state = self.extract_info(
            driver = driver, 
            element_id='txtState',
            field_name= 'State'
            )
        zip_code = self.extract_info(
            driver = driver, 
            element_id='txtZipCode',
            field_name= 'Zip')
        demographic_address = street + unit + ', ' + city + ', ' + state + ' ' + zip_code

            # Phone number
        home_phone = self.extract_info(
            driver = driver, 
            element_id='txtHomePhone',
            field_name='Home Telephone')
            # cell = self.extract_info(driver, 'txtCellPhone')

            # Gender
        gender = self.dropDown_extract(driver, 'cboGender')


            # -------- Lab tab ------------------
            # Navigate to lab tab 
        logging.info('Switching to Lab Tab')
        lab_tab_id = 'ctl37_btnSupplementalTabSYS'
        lab_tab = driver.find_element(By.ID, lab_tab_id)
        lab_tab.click()
        #time.sleep(5)
        '''
        # skip if there is not any info on lab tab
        x = input('Is there info on lab tab (y/n)?...')
        if x in ('y', 'yes', 'Y', 'Yes'): 
            x = None
        elif x in ('n', 'No', 'no', 'N'): 
            x = None
            self.nav2IMM(driver)
            return 'skip'
        else: pass
        # end 
         '''
        logging.info('Grabbing information from Lab Tab')
            # accession number
        acc_num = self.extract_info(
            driver=driver, 
            element_id='_-11_ctl03_dgLabInfo_ctl02_txtAccNum', 
            field_name='Accession Number'
            )
            # Specimen Collected Date
        specimen_collect_date = self.extract_info(
            driver = driver,
            element_id="_-11_ctl03_dgLabInfo_ctl02_txtSpecCollDate",
            field_name='Specimen Collection Date'
            )
            # Specimen Received Date 
        specimen_received_date = self.extract_info(
            driver=driver,
            element_id= "_-11_ctl03_dgLabInfo_ctl02_txtSpecReceDate",
            field_name='Specimen Receive Date'
            )
            # Specimen Source
        specimen_source = self.extract_info(
            driver = driver,
            element_id="_-11_ctl03_dgLabInfo_ctl02_txtSpecimenSourceText",
            field_name='Specimen Source'
            )
            # Resulted Test
        resulted_test = self.extract_info(
            driver = driver,
            element_id="_-11_ctl03_dgLabInfo_ctl02_txtResultedTest",  # might need to change id back to put 'L' at the end
            field_name='Resulted Test', 
            xpath='//*[@id="_-11_ctl03_dgLabInfo_ctl02_txtResultedTestL"]'
            )
            # result
        result = self.extract_info(
            driver = driver,
            element_id="_-11_ctl03_dgLabInfo_ctl02_txtResult", # might need to change id back to put 'L' at the end
            field_name='Result',
            xpath='//*[@id="_-11_ctl03_dgLabInfo_ctl02_txtResult"]'
            )
            # resulted organism 
        result_organism = self.extract_info(driver, 
            element_id="_-11_ctl03_dgLabInfo_ctl02_txtResultedOrganism",
            field_name= 'Resulted Organism',
            xpath= '//*[@id="_-11_ctl03_dgLabInfo_ctl02_txtResultedOrganismL"]'
            )
            # units
        units = self.extract_info(
            driver = driver, 
            element_id='_-11_ctl03_dgLabInfo_ctl02_txtUnits',
            field_name='Units'
            )
            # Reference Range
        ref_range = self.extract_info(
            driver = driver,
            element_id='_-11_ctl03_dgLabInfo_ctl02_txtReferencerange',
            field_name='Reference Range'
            )
            # Result date
        result_date = self.extract_info(
            driver = driver, 
            element_id='_-11_ctl03_dgLabInfo_ctl02_txtResultDate',
            field_name= 'Result Date'
            )
            # Performing facility ID, 
        perform_facility_id = self.extract_info(
            driver,
            element_id='_-11_ctl03_dgLabInfo_ctl02_txtPerformingFacilityID',
            field_name='Performing Facility ID'
            )
            # Abnormal Flag
        ab_flag = self.dropDown_extract(
            driver=driver,
            select_id='_-11_ctl03_dgLabInfo_ctl02_ddlAbnormalFlag',
            menu_id='_-11_ctl03_dgLabInfo_ctl02_ddlAbnormalFlag'
            )
            # Observation Ressults
        ob_results = self.dropDown_extract(
            driver=driver,
            select_id='_-11_ctl03_dgLabInfo_ctl02_ddlObservationResultStat',
            menu_id= '/html/body/form/div[2]/div/div/table[2]/tbody/tr[5]/td/table/tbody/tr/td/div/div/table/tbody/tr/td/table/tbody/tr[1]/td/div[1]/table[2]/tbody/tr[10]/td/table/tbody/tr[4]/td[1]/div[1]/select')

            # Provider name
        provider_name = self.extract_info(
            driver=driver,
            element_id= '_-11_ctl03_dgLabInfo_ctl02_txtProviderName',
            field_name='Provider Name'
            )
            # Order callback number
        provider_number = self.extract_info(
            driver = driver,
            element_id='_-11_ctl03_dgLabInfo_ctl02_txtProviderCallBack',
            field_name='Provider Number'
            )
            # Provider address 
        list_of_addressIDs = [
                ['_-11_ctl03_dgLabInfo_ctl02_txtProviderAddress', 'Address'],
                ['_-11_ctl03_dgLabInfo_ctl02_txtProviderCity', 'City'],
                ['_-11_ctl03_dgLabInfo_ctl02_txtProviderState','State'],
                ['_-11_ctl03_dgLabInfo_ctl02_txtFacilityZip', 'Zip']
                ]
        provider_address = self.address(driver, list_of_addressIDs)

            # Facility name 
        facility_name = self.extract_info(
            driver = driver,
            element_id='_-11_ctl03_dgLabInfo_ctl02_txtFacilityName',
            field_name= 'FacilityName')
            
            # Facility address
        facility_address_ids = [
                ['_-11_ctl03_dgLabInfo_ctl02_txtFacilityAddress', 'Address'],
                ['_-11_ctl03_dgLabInfo_ctl02_txtFacilityCity', 'City'],
                ['_-11_ctl03_dgLabInfo_ctl02_txtFacilityState', 'State'],
                ['_-11_ctl03_dgLabInfo_ctl02_txtFacilityZip', 'Zip']
            ]
        facility_address = self.address(driver, facility_address_ids)

            # Facility phone number
        facility_phone = self.extract_info(
            driver=driver,
            element_id='_-11_ctl03_dgLabInfo_ctl02_txtFacilityPhone',
            field_name='Facility Phone'
            )
            # have to click cancel after I am done with getting information from both tabs
        logging.info('Finished grabbing information from WebCMR website')
        cancel_id = 'btnCancel'
        cancel_btn = driver.find_element(By.ID, cancel_id)
        cancel_btn.click()            

            # storing all of this information in a list 
            
        webCMR_values = [
                name,
                dob,
                race,
                ethnicity,
                demographic_address,
                home_phone,
                gender,
                acc_num,
                specimen_collect_date,
                specimen_received_date,
                specimen_source, 
                resulted_test,
                result,
                result_organism,
                units,
                ref_range,
                result_date,
                ab_flag,
                ob_results,
                provider_name,
                provider_address,
                provider_number,
                perform_facility_id,
                facility_name,
                facility_address,
                facility_phone
            ]
        webCMR_indicies = [
                'Name',
                'DOB',
                'Race',
                'Ethnicity',
                'Demographic Address',
                'Phone',
                'Gender',
                'Accession Number',
                'Specimen Collected Date',
                'Specimen Received Date',
                'Specimen Source',
                'Resulted Test',
                'Result',
                'Resulted Organism',
                'Units',
                'Reference Range',
                'Result Date',
                'Abnormal Flag',
                'Observation Results',
                'Provider Name',
                'Provider Address',
                'Provider Number',
                'Performing Facility ID',
                'Facility Name',
                'Facility Address',
                'Facility Phone'
            ]
            # test_name, acc_num = df['DILR_ResultedTest'][i], df['DILR_AccessionNumber'][i]
            
            # Replace special characters and spaces with "_"
            # file_name = re.sub(r'[^\w\s]+', '_', f'{test_name}_AccNum_{acc_num}_IncidentID_{di_num}')

            # curr_dir = os.getcwd()
            # webCMR_df = pd.DataFrame(data=webCMR_values, columns=['TSTWebCMR Data'], index=webCMR_indicies)
            # webCMR_df.to_excel(f'{curr_dir}/{file_name}.xlsx')

            # navigating to IMM menu after DI searches
        _ = self.nav2IMM(driver)
        return (webCMR_values,webCMR_indicies)

    def address(self, driver, ids : list):
        prov_st = self.extract_info(
            driver= driver,
            element_id=ids[0][0],
            field_name=ids[0][1]
            )
        prov_city = self.extract_info(
            driver=driver,
            element_id=ids[1][0],
            field_name=ids[1][1]
            )
        prov_state = self.extract_info(
            driver=driver,
            element_id=ids[2][0],
            field_name=ids[2][1]
            )
        prov_zip = self.extract_info(
            driver = driver,
            element_id=ids[3][0],
            field_name=ids[3][1]
            )
        provider_address = prov_st + ', ' + prov_city +', ' + prov_state + ' ' + prov_zip
        return provider_address

    def dropDown_extract(self, driver,  select_id, menu_id=None):
        '''
        Alternative method for getting info from dropdown menu
        '''
        try:
            dropdown_element = driver.find_element(By.ID, select_id)
            selected_option = Select(dropdown_element).first_selected_option
        except NoSuchElementException:
            if menu_id:
                try:
                    dropdown_element = driver.find_element(By.XPATH, menu_id)
                    selected_option = Select(dropdown_element).first_selected_option
                except NoSuchElementException:
                    text = 'CODE ERROR....element not found but is there'
                    return text
            else:
                raise NoSuchElementException(f"Menu with ID {menu_id} and Select with {select_id} not found.")
        text = selected_option.text
        return text
    def extract_info(self, driver, element_id, xpath=None, field_name=None ):
        
        '''
        Function is meant to extract information from text boxes, 
        using the elements ID as a identifier. If element is not found 
        by ID, function will attempt to find it by XPath.
        '''
        try: 
            element_btn = driver.find_element(By.ID, element_id)
        except NoSuchElementException:
            if xpath:
                element_btn = driver.find_element(By.XPATH, xpath)
            elif field_name: 
                # field name is a fail-safe to find element if the id by itself is not working
                # typing in the filed name will remind me which section to look at if it breaks
                element_btn = driver.find_element(By.XPATH, f"//*[contains(@id, {element_id})]")
                print(f'Except Block: {field_name}')
            else:
                raise NoSuchElementException(f"Element with ID {element_id} and XPath {xpath} not found.")
        
        value = element_btn.get_attribute('value')
        
        return value