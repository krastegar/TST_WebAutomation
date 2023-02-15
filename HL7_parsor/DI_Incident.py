import time
import os
import pandas as pd
import re

from IMM import IMM
from SetUp import SetUp
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

class DI_Search(IMM, SetUp): 

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
    
    def disease_search(self): 
        # Login into home page
        # driver = self.login() 
        driver = self.imm_export()

        # Getting Disease Queries and Hl7 accession #
        df = self.export_df()
        di_num_list = list(df['DILR_IncidentID'])

        # click home button 
        home_btn = driver.find_element(By.ID, 'FragTop1_lbtnHome')
        home_btn.click()
        webCMR_values, webCMR_indicies = None, None # bounding these variables
        for i in range(len(di_num_list)):
            di_num = int(di_num_list[i])
            
            # navigating to disease incidents tab
            DI_tab = driver.find_element(By.ID, 'ibtnTabDisInc')
            DI_tab.click()
            
            # Searching for specific disease incident 
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
            
            # Name fields
            last_name = self.extract_info(driver, 'txtLastName')
            first_name = self.extract_info(driver, 'txtFirstName')
            name = first_name + ' ' + last_name

            # DOB
            dob = self.extract_info(driver, 'txtDOB')

            # Reported Race
            race = self.extract_info(driver, 'txtReportedRace')

            # Ethnicity
            ethnicity = self.dropDown_extract(driver, 'cboEthnicity')

            # Address
            street = self.extract_info(driver, 'txtAddress')
            unit = self.extract_info(driver, 'txtApartment')
            city = self.extract_info(driver, 'txtCity')
            state = self.extract_info(driver, 'txtState')
            zip_code = self.extract_info(driver, 'txtZipCode' )
            demographic_address = street + unit + ', ' + city + ', ' + state + ' ' + zip_code

            # Phone number
            home_phone = self.extract_info(driver, 'txtHomePhone')
            # cell = self.extract_info(driver, 'txtCellPhone')

            # Gender
            gender = self.dropDown_extract(driver, 'cboGender')


            # -------- Lab tab ------------------
            # Navigate to lab tab 
            lab_tab_id = 'ctl37_btnSupplementalTabSYS'
            lab_tab = driver.find_element(By.ID, lab_tab_id)
            lab_tab.click()

            # accession number
            acc_num = self.extract_info(driver, 
                                '_-11_ctl03_dgLabInfo_ctl02_txtAccNum')

            # Specimen Collected Date
            specimen_collect_date = self.extract_info(driver,
                                "_-11_ctl03_dgLabInfo_ctl02_txtSpecCollDate")
            # Specimen Received Date 
            specimen_received_date = self.extract_info(driver,
                                "_-11_ctl03_dgLabInfo_ctl02_txtSpecReceDate")
            
            # Specimen Source
            specimen_source = self.extract_info(driver,
                                "_-11_ctl03_dgLabInfo_ctl02_txtSpecimenSourceText")
            
            # Resulted Test
            resulted_test = self.extract_info(driver,
                                "_-11_ctl03_dgLabInfo_ctl02_txtResultedTestL")
            
            # result
            result = self.extract_info(driver,
                                "_-11_ctl03_dgLabInfo_ctl02_txtResult")
            
            # resulted organism 
            result_organism = self.extract_info(driver, 
                                "_-11_ctl03_dgLabInfo_ctl02_txtResultedOrganism")

            # units
            units = self.extract_info(driver, 
                                '_-11_ctl03_dgLabInfo_ctl02_txtUnits')
            
            # Reference Range
            ref_range = self.extract_info(driver,
                                '_-11_ctl03_dgLabInfo_ctl02_txtReferencerange')

            # Result date
            result_date = self.extract_info(driver, 
                                '_-11_ctl03_dgLabInfo_ctl02_txtResultDate')
            
            # Performing facility ID, 
            perform_facility_id = self.extract_info(driver,
                                '_-11_ctl03_dgLabInfo_ctl02_txtPerformingFacilityID')
            

            # Abnormal Flag
            ab_flag = self.dropDown_extract(driver,
                                '_-11_ctl03_dgLabInfo_ctl02_ddlAbnormalFlag')
            
            # Observation Ressults
            ob_results = self.dropDown_extract(driver,
                                '_-11_ctl03_dgLabInfo_ctl02_ddlObservationResultStat')

            # Provider name
            provider_name = self.extract_info(driver,
                                '_-11_ctl03_dgLabInfo_ctl02_txtProviderName')

            # Order callback number
            provider_number = self.extract_info(driver,
                                '_-11_ctl03_dgLabInfo_ctl02_txtProviderCallBack')

            # Provider address 
            list_of_addressIDs = [
                '_-11_ctl03_dgLabInfo_ctl02_txtProviderAddress',
                '_-11_ctl03_dgLabInfo_ctl02_txtProviderCity',
                '_-11_ctl03_dgLabInfo_ctl02_txtProviderState',
                '_-11_ctl03_dgLabInfo_ctl02_txtFacilityZip'
                ]
            provider_address = self.address(driver, list_of_addressIDs)

            # Facility name 
            facility_name = self.extract_info(driver,
                                '_-11_ctl03_dgLabInfo_ctl02_txtFacilityName')
            
            # Facility address
            facility_address_ids = [
                '_-11_ctl03_dgLabInfo_ctl02_txtFacilityAddress',
                '_-11_ctl03_dgLabInfo_ctl02_txtFacilityCity',
                '_-11_ctl03_dgLabInfo_ctl02_txtFacilityState',
                '_-11_ctl03_dgLabInfo_ctl02_txtFacilityZip'
            ]
            facility_address = self.address(driver, facility_address_ids)

            # Facility phone number
            facility_phone = self.extract_info(driver,
                                '_-11_ctl03_dgLabInfo_ctl02_txtFacilityPhone')
            

            # have to click cancel after I am done with getting information from both tabs
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
            test_name, acc_num = df['DILR_ResultedTest'][i], df['DILR_AccessionNumber'][i]
            
            # Replace special characters and spaces with "_"
            file_name = re.sub(r'[^\w\s]+', '_', f'{test_name}_AccNum_{acc_num}_IncidentID_{di_num}')

            curr_dir = os.getcwd()
            webCMR_df = pd.DataFrame(data=webCMR_values, columns=['TSTWebCMR Data'], index=webCMR_indicies)
            webCMR_df.to_excel(f'{curr_dir}/{file_name}.xlsx')

            # navigating to IMM menu after DI searches
        _ = self.nav2IMM(driver)
        
        # time.sleep(8)

        return driver, df, webCMR_values, webCMR_indicies

    def address(self, driver, ids : list):
        prov_st = self.extract_info(driver,
                                ids[0])
        prov_city = self.extract_info(driver,
                               ids[1])
        prov_state = self.extract_info(driver,
                                ids[2])
        prov_zip = self.extract_info(driver,
                                ids[3])
        provider_address = prov_st + ', ' + prov_city +', ' + prov_state + ' ' + prov_zip
        return provider_address

    def dropDown_extract(self, driver, menu_id):
        '''
        Alternative method for getting info from dropdown menu
        '''
        dropdown_element = driver.find_element(By.ID, menu_id)
        selected_option = Select(dropdown_element).first_selected_option
        text = selected_option.text
        return text

    def extract_info(self, driver, element_id):
        '''
        Function is meant to extract information from text boxes, 
        using the elements ID as a identifier. 
        '''
        element_btn = driver.find_element(By.ID, element_id)
        value = element_btn.get_attribute('value')
        return value