import os
import pandas as pd 
import re
import numpy as np

from selenium.webdriver.common.by import By
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from DI_Incident import DI_Search
from IMM import IMM


class HL7_extraction(DI_Search, IMM):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        return
    
    def hl7_copy(self):
        '''
        In the final version I wll have disease_search in this method 
        '''
        
        # Going into IMM Menu and Downloading excel sheet 
        driver = self.imm_export()

        # Getting Disease Queries and Hl7 accession # from excel sheet
        df = self.export_df()
        
        # Navigate to IMM menu after export
        _ = self.nav2IMM(driver)

        acc_num_list = list(df['DILR_AccessionNumber'])
        di_num_list = list(df['DILR_IncidentID'])
        # Resulted test list for searches 
        test_list = list(df['DILR_ResultedTest'])
        wait = WebDriverWait(driver, 10)
        for i in range(len(acc_num_list)):
            print(f'\nITERATION #: {i}')
            try: 
                try: 
                    # Searching IMM for specific accession numbers 
                    # Going to have to put this in a for loop 
                    acc_num = int(acc_num_list[i])
                    resultTest = str(test_list[i])
                    di_num = int(di_num_list[i])
                    acc_box, test_search = self.acc_test_search(wait, acc_num, resultTest)

                    # Copy and filter hl7 results
                    hl7_values, hl7_section_labels = self.data_wrangling(driver, resultTest, acc_num)
                    
                    # get webCMR values
                    webCMR_values, webCMR_indicies = self.webTST_scrape(driver, di_num_list, i)
                    print(
                        f'''
                        printing lengths: 
                        webCMR_values = {len(webCMR_values)}
                        webCMR_indicies = {len(webCMR_indicies)}
                        hl7_values = {len(hl7_values)}, 
                        hl7_sections = {len(hl7_section_labels)}
                        '''
                    )
                    # make summary df

                    webCMR_hl7_df = pd.DataFrame(
                        data=np.array([webCMR_values,hl7_values, hl7_section_labels]).T, 
                        columns=['TSTWebCMR Data', 'HL7 Data', 'HL7 Section'], 
                        index=webCMR_indicies
                        )

                    # make summarry report 
                    report = self.hl7_report(resultTest, acc_num, di_num,webCMR_hl7_df)
                    
                except StaleElementReferenceException as e:

                    # seaching accession and resulted test again 
                    acc_num = int(acc_num_list[i])
                    resultTest = str(test_list[i])
                    di_num = int(di_num_list[i])
                    acc_box, test_search = self.acc_test_search(wait, acc_num, resultTest)
                    
                    # Copy and filter hl7 results
                    hl7_values, hl7_section_labels = self.data_wrangling(driver, resultTest, acc_num)

                    webCMR_values, webCMR_indicies = self.webTST_scrape(driver, di_num_list, i)
                    
                    # make summary df

                    webCMR_hl7_df = pd.DataFrame(
                        data=np.array([webCMR_values,hl7_values, hl7_section_labels]).T, 
                        columns=['TSTWebCMR Data', 'HL7 Data', 'HL7 Section'], 
                        index=webCMR_indicies
                        )
                    # produce report 
                    report = self.hl7_report(resultTest, acc_num, webCMR_hl7_df)
            except StaleElementReferenceException as e:
                continue
        home_directory = os.path.expanduser( '~' )
        new_dir = f'{home_directory}/Desktop/hl7_test'
        query_df = df.to_excel(f'{new_dir}/query.xlsx')
    

    def acc_test_search(self, wait, acc_num, resultTest):
        acc_box = wait.until(EC.presence_of_element_located((By.ID, 'txtAccession')))
        acc_box.clear()
        acc_box.send_keys(str(acc_num))
        test_search = wait.until(EC.presence_of_element_located((By.ID, 'txtResultedTest')))
        test_search.clear()
        test_search.send_keys(resultTest)
        search_id = 'ibtnSearch'
        search_btn = wait.until(EC.presence_of_element_located((By.ID, search_id)))
        search_btn.click()
        return acc_box,test_search

    def data_wrangling(self, driver, resultTest, acc_num): 
        # Get HL7 from contents area after IMM search 
        table = driver.find_element(By.ID, "divContentsArea").text
        hl7_sections = table.split('\n')
        components = [section.split("|") for section in hl7_sections]
        df = pd.DataFrame(components)
        
        obx_indx, obr_indx, spm_indx = [], [], []
        for index, section in enumerate(list(df[0].values)):
            if section == 'OBR':
                obr_indx.append(index)
            elif section == 'OBX': 
                obx_indx.append(index)
            elif section == 'SPM': 
                spm_indx.append(index)
            else: continue 
        
        #print(f'# of OBX: {len(obx_indx)} \n # of OBR segments: {len(obr_indx)} \n # of SPM segments:  {len(spm_indx)}')
        
        try:
            assert len(obr_indx) == len(obx_indx) == len(spm_indx)
            # Going to check which OBX test result matches the imported test result
            column = 3 # OBX 3 is the column that contains resulted test 
            key = self.match_obx(resultTest, df, obx_indx, column)
            pos = obx_indx.index(key)
            df_idx = [obr_indx[pos], obx_indx[pos], spm_indx[pos]]
        except AssertionError:
            column = 3
            key = self.match_obx(resultTest, df, obx_indx, column)
            pos = obx_indx.index(key)
            if len(obr_indx) < len(obx_indx): 
                df_idx = [obr_indx[0], obx_indx[pos], spm_indx[pos]]
            elif len(spm_indx) < len(obx_indx): 
                df_idx = [obr_indx[pos], obx_indx[pos], spm_indx[0]]
            else: pass 
        
        sections : list = list(df[0].values)
        # not sure if there is going to be more than one ORC section, 
        # Marjorie seemed hesistant around this 
        obr_cnt = 0
        for sect in sections: 
            if sect == 'ORC': 
                obr_cnt+=1
                df_idx.append(sections.index(sect))
                if obr_cnt > 1: 
                    raise ValueError('Duplicate ORC Sections')
            elif sect == 'PID': 
                df_idx.append(sections.index(sect))
            else: pass 
        
        # making pandas df with only rows related to ResultedTest of interest
        oneResult_df = df.iloc[df_idx]
        oneResult_df.set_index(0, inplace=True)
        
        # print(f'name column: {name}')
        # going to extract values from hl7
        name = oneResult_df.loc['PID', 5]
        dob = oneResult_df.loc['PID', 7]
        race = oneResult_df.loc['PID', 10]
        ethnicity = oneResult_df.loc['PID', 22]
        address = oneResult_df.loc['PID', 11]
        phone_number = oneResult_df.loc['PID', 13]
        gender = oneResult_df.loc['PID', 8]
        accession = self.get_accession(oneResult_df.loc['SPM', 2])
        specimen_collect = self.check_specimenCollect(oneResult_df)
        specimen_receive = oneResult_df.loc['SPM', 18]
        specimen_source = oneResult_df.loc['SPM', 4]
        resulted_test = oneResult_df.loc['OBX', 3]
        result = oneResult_df.loc['OBX', 5]
        resulted_organism = 'If there is one it is in Result section'
        units = oneResult_df.loc['OBX', 6]
        ref_range = oneResult_df.loc['OBX', 7]
        result_date = oneResult_df.loc['OBX', 19]
        performing_facility_ID = oneResult_df.loc['OBX', 23]
        ab_flag = oneResult_df.loc['OBX', 8]
        ob_results = oneResult_df.loc['OBX', 11]
        provider_name = oneResult_df.loc['ORC', 12]
        provider_phone = oneResult_df.loc['ORC', 14]
        provider_address = oneResult_df.loc['ORC', 24]
        facility_name = oneResult_df.loc['ORC', 21]
        facility_address = oneResult_df.loc['ORC', 22]
        facility_phone = oneResult_df.loc['ORC', 23]

        hl7_section_label = [
            'PID-5',
            'PID-7',
            'PID-10',
            'PID-22',
            'PID-11',
            'PID-13',
            'PID-8',
            'SPM-2',
            'SPM-17, OBR-7, OBX-14',
            'SPM-18',
            'SPM-4', 
            'OBX-3',
            'OBX-5', 
            'N/A',
            'OBX-6',
            'OBX-7',
            'OBX-19',
            'OBX-23',
            'OBX-8',
            'OBX-11',
            'ORC-12',
            'ORC-14',
            'ORC-24',
            'ORC-21',
            'ORC-22',
            'ORC-23'
        ]
        
        hl7_values = [
            name,
            dob,
            race,
            ethnicity,
            address,
            phone_number,
            gender,
            accession,
            specimen_collect,
            specimen_receive,
            specimen_source,
            resulted_test,
            result,
            resulted_organism,
            units,
            ref_range,
            result_date,
            ab_flag,
            ob_results,
            provider_name,
            provider_address,
            provider_phone,
            performing_facility_ID,
            facility_name, 
            facility_address,
            facility_phone
        ]
        # Need to return home so that we can go to disease tab next 
        home_btn = driver.find_element(By.ID, 'FragTop1_lbtnHome')
        home_btn.click()

        return hl7_values, hl7_section_label
    
    def get_accession(self, string ):
        'Return first part of SPM 2 that is seen before accession #'
        return string.split('&', 1)[0]

    def check_specimenCollect(self, df):
        spm, obx, obr = df.loc['SPM', 17], df.loc['OBX', 14], df.loc['OBR', 7]
        if spm == obx == obr:
            sp_collect = spm
        elif spm == obx:
            sp_collect = f'SPM-17 & OBX-14 is {obx}, while OBR-7 is {obr}'
        elif spm == obr:
            sp_collect = f'SPM-17 & OBR-7 is {spm}, while OBR-7 is {obx}'
        elif obx == obr:
            sp_collect = f'OBX-14 & OBR-7 is {obr}, while SPM-17 is {spm}'
        else:
            sp_collect=f"""
            All variables are different.
            SPM-17: {spm}
            OBX-14: {obx}
            OBR-7: {obr}
            """
        return sp_collect

    def hl7_report(self, resultTest, acc_num, di_num,df):
        home_directory = os.path.expanduser( '~' )
        new_dir = f'{home_directory}/Desktop/hl7_test/'
        try:
            mkdir = os.mkdir(new_dir)
        except FileExistsError:
            pass 
        file_name = re.sub(r'[^\w\s]+', '_', resultTest)
        file_dir = f'{new_dir}/{file_name}'
        df.to_excel(f'{file_dir}_ACCnum_{acc_num}_DInum_{di_num}.xlsx')

    def match_obx(self, resultTest, df, obx_indx, column):
        key : int = 0
        for row_index in obx_indx: 
            obx_result : str = str(df.iloc[row_index, column])
            if resultTest in obx_result:
                # print(f'Found the Index: {row_index}') 
                key = int(row_index)
        return key