import os
import pandas as pd 
import re
import numpy as np
import logging
import time
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
        logging.info('Exporting DataFrame')
        # Going into IMM Menu and Downloading excel sheet 
        driver = self.imm_export()

        logging.info("Random Sampling from Export....grabing unique test com")
        # Getting Disease Queries and Hl7 accession # from excel sheet
        df = self.export_df()
        
        # Navigate to IMM menu after export
        _ = self.nav2IMM(driver) 
        
        logging.info("Doing Searches for HL7 and Disease Incident")
        acc_num_list = list(df['DILR_AccessionNumber'])
        di_num_list = list(df['DILR_IncidentID'])
        # Resulted test list for searches 
        test_list = list(df['DILR_ResultedTest'])

        for i in range(len(acc_num_list)):
            print(f'\nITERATION #: {i}')
            logging.info(f'\nITERATION #: {i}')
            try: 
                try: 
                    # Searching IMM for specific accession numbers 
                    # Going to have to put this in a for loop 
                    acc_num = str(acc_num_list[i])
                    resultTest = str(test_list[i])
                    di_num = int(di_num_list[i]) 

                    logging.info("Inputting ResultedTest and Accession numbers in IMM search boxes")

                    acc_box, test_search = self.acc_test_search(driver, acc_num, resultTest)

                    # Copy and filter hl7 results
                    hl7_values, hl7_section_labels = self.data_wrangling(driver, resultTest, acc_num)
                    
                    # get webCMR values
                    logging.info("Performing Webscraping")
                    webCMR_values, webCMR_indicies = self.webTST_scrape(driver, di_num_list, i)
                    '''
                    web_values = self.webTST_scrape(driver, di_num_list, i)
                    if web_values == 'skip':
                        continue
                    else: 
                        webCMR_values, webCMR_indicies = web_values
                    '''

                    # make summary df
                    webCMR_hl7_df = pd.DataFrame(
                        data=np.array([webCMR_values,hl7_values, hl7_section_labels]).T, 
                        columns=['TSTWebCMR Data', 'HL7 Data', 'HL7 Section'], 
                        index=webCMR_indicies
                        )

                    # make summarry report 
                    logging.info('Making summary Report folders')
                    report = self.hl7_report(resultTest, acc_num, di_num,webCMR_hl7_df)
                    
                except StaleElementReferenceException as e:
                    # Searching IMM for specific accession numbers 
                    # Going to have to put this in a for loop 
                    acc_num = int(acc_num_list[i])
                    resultTest = str(test_list[i])
                    di_num = int(di_num_list[i])

                    logging.info("Inputting ResultedTest and Accession numbers in IMM search boxes")
                    acc_box, test_search = self.acc_test_search(driver, acc_num, resultTest)

                    # Copy and filter hl7 results
                    hl7_values, hl7_section_labels = self.data_wrangling(driver, resultTest, acc_num)
                    
                    # get webCMR values
                    logging.info("Performing Webscraping")
                    webCMR_values, webCMR_indicies = self.webTST_scrape(driver, di_num_list, i)
                    '''
                    web_values = self.webTST_scrape(driver, di_num_list, i)
                    if web_values == 'skip':
                        continue
                    else: 
                        webCMR_values, webCMR_indicies = web_values
                    '''

                    # make summary df
                    webCMR_hl7_df = pd.DataFrame(
                        data=np.array([webCMR_values,hl7_values, hl7_section_labels]).T, 
                        columns=['TSTWebCMR Data', 'HL7 Data', 'HL7 Section'], 
                        index=webCMR_indicies
                        )

                    # make summarry report 
                    logging.info('Making summary Report folders')
                    report = self.hl7_report(resultTest, acc_num, di_num,webCMR_hl7_df)
            except StaleElementReferenceException as e:
                continue
        #home_directory = os.path.expanduser( '~' )
        lab_name = re.sub(r'[^\w\s]+', '_',self.lab)
        new_dir = f'./{lab_name}/'
        query_df = df.to_excel(f'{new_dir}/query_summary.xlsx')
    

    def acc_test_search(self, driver, acc_num, resultTest):
        #acc_box = wait.until(EC.presence_of_element_located((By.ID, 'txtAccession')))
        acc_box = self.multiFind(
            driver=driver,
            element_id= 'txtAccession',
            xpath='/html/body/form/div[3]/div/div/table[3]/tbody/tr[2]/td/table/tbody/tr[1]/td[8]/input'
        )
        acc_box.clear()
        acc_box.send_keys(str(acc_num))
        #test_search = wait.until(EC.presence_of_element_located((By.ID, 'txtResultedTest')))
        test_search = self.multiFind(
            driver=driver,
            element_id= 'txtResultedTest',
            xpath='/html/body/form/div[3]/div/div/table[3]/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/input'
        )
        test_search.clear()
        test_search.send_keys(resultTest)
        search_id = 'ibtnSearch'
        #search_btn = wait.until(EC.presence_of_element_located((By.ID, search_id)))
        search_btn = self.multiFind(
            driver=driver,
            element_id= search_id,
            xpath='/html/body/form/div[3]/div/div/table[3]/tbody/tr[2]/td/table/tbody/tr[4]/td/div/input[1]'
        )
        search_btn.click()
        return acc_box,test_search

    def data_wrangling(self, driver, resultTest, acc_num): 
        
        df, obx_indx, obr_indx, spm_indx = self.hl7_redoSearch(driver, resultTest, acc_num)
        
        print(f'# of OBX: {len(obx_indx)} \n # of OBR segments: {len(obr_indx)} \n # of SPM segments:  {len(spm_indx)}')
        #print(resultTest)
        try:
            assert len(obr_indx) == len(obx_indx) == len(spm_indx)
            
            # Going to check which OBX test result matches the imported test result
            column = 3 # OBX 3 is the column that contains resulted test 
            key = self.match_obx(resultTest, df, obx_indx, column)  
            
            if key == 0: # there was an issue with a search if this value is 0
                # going to redo search if the initial search does not provide HL7 message
                logging.info('No OBX section found....redoing search')
                print('No OBX section found....redoing search')

                # issue is usually caused by brackets being in the resultTest name 
                resultTest = re.sub(r'\[.*?\]', '', resultTest) # removes brackets and anything inside of them

                # redoing initial hl7 search 
                df, obx_indx, obr_indx, spm_indx = self.hl7_redoSearch(driver, resultTest, acc_num)
                key = self.match_obx(resultTest, df, obx_indx, column)
                
                #print(f'Key after match_box(): {key}')
                # finding the obr, obx, and spm sections that are corresponding to incoming message monitor 
                # ResultedTest value
                pos = obx_indx.index(key)
                df_idx = [obr_indx[pos], obx_indx[pos], spm_indx[pos]]


            # finding the obr, obx, and spm sections that are corresponding 
            # to incoming message monitor ResultedTest value
            pos = obx_indx.index(key)
            df_idx = [obr_indx[pos], obx_indx[pos], spm_indx[pos]]


        except AssertionError:
            column = 3
            key = self.match_obx(resultTest, df, obx_indx, column)
            pos = obx_indx.index(key)
            if len(spm_indx) < len(obx_indx) and len(obr_indx) < len(obx_indx): 
                df_idx = [obr_indx[0], obx_indx[pos], spm_indx[0]]
            elif len(spm_indx) < len(obx_indx) and len(obr_indx) == len(obx_indx): 
                df_idx = [obr_indx[pos], obx_indx[pos], spm_indx[0]]
            elif len(obr_indx) < len(obx_indx) and len(spm_indx)==len(obx_indx): 
                df_idx = [obr_indx[0], obx_indx[pos], spm_indx[pos]]
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
        self.go_home(driver)

        return hl7_values, hl7_section_label

    def hl7_redoSearch(self, driver, resultTest, acc_num):
        '''
        Function is meant to redo a search for hl7 message if there are no results 
        searching for the resultedTest and accession number in Incoming Message monitor page. 
        '''
        self.nav2IMM(driver=driver)
        time.sleep(1)
        self.acc_test_search(resultTest=resultTest,
                                acc_num=acc_num,
                                driver=driver)
        # Get HL7 from contents area after IMM search 
        table = driver.find_element(By.ID, "divContentsArea").text
        hl7_sections = table.split('\n')
        components = [section.split("|") for section in hl7_sections]
        df = pd.DataFrame(components)
        obx_indx, obr_indx, spm_indx = [], [], []
        for index, section in enumerate(list(df[0].values)):
            if section.upper() == 'OBR':
                obr_indx.append(index)
            elif section.upper() == 'OBX': 
                obx_indx.append(index)
            elif section.upper() == 'SPM': 
                spm_indx.append(index)
            else: continue 
        return df, obx_indx, obr_indx, spm_indx
    
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
            sp_collect = f'SPM-17 & OBR-7 is {spm}, while OBX-14 is {obx}'
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
        lab_name = re.sub(r'[^\w\s]+', '_',self.lab)
        new_dir = f'./{lab_name}/'
        try:
            mkdir = os.mkdir(new_dir)
        except FileExistsError:
            pass 
        file_name = re.sub(r'[^\w\s]+', '_', resultTest)
        file_dir = f'{new_dir}/{file_name}'
        df.to_excel(f'{file_dir}_ACCnum_{acc_num}_DInum_{int(di_num)}.xlsx')

    def match_obx(self, resultTest : str, df : pd.DataFrame, obx_indx : int, column : int):
        key = 0
        for row_index in obx_indx: 
            obx_result : str = str(df.iloc[row_index, column])
            if resultTest in obx_result:
                #print(f'\nFound the Index: {row_index}\n') 
                key = int(row_index)
                print(f'Return Key in if-else: {key}')
                break
        print(f'Return value: {key}')
        return key