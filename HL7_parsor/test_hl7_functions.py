import time
import pandas as pd 
from DI_Incident import DI_Search
from IMM import IMM
from HL7 import HL7_extraction

def test_single_accession(): 
    '''
    The purpose of this function is to test specific ELR cases instead of running through
    the entire process of the program which takes up a lot of time. Going to start at the 
    part of the process after downloading incoming message monitor file 
    '''

    # Reading already existing IMM excel file from downloads folder 
    # and inspecting variables in IMM class's export_df() method. 
    imm = HL7_extraction(
            username = 'krastegar',
            paswrd = 'Hamid&Mahasty3',
            FromDate= '02/13/2023',
            ToDate= '02/24/2023',
            LabName= "Palomar Health Downtown Campus Laboratory NEW"
            )
    
    # testing random sampling function
    # This is the roadmap of the entire program.
    imm_df=imm.export_df()
    
    # pulling specific row where we are seeing issues in the program 
    issue_row = imm_df.iloc[73]
    
    # grabbing values for searching on TST-WebCMR 
    acc_num = str(issue_row['DILR_AccessionNumber'])
    resultTest = str(issue_row['DILR_ResultedTest'])
    #di_num = str(issue_row['DILR_IncidentID']) # GISAID sequence accession number Nom (Isol) [ID]

    # calling functions to test the changes
    driver = imm.login()
    imm.nav2IMM(driver=driver)
    imm.acc_test_search(driver, acc_num, resultTest)
    parse_hl7=imm.data_wrangling(driver, resultTest, acc_num)



    return 

if __name__=="__main__":
    imm_df = test_single_accession()
    None

