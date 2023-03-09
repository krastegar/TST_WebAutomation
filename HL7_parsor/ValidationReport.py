import sys
import os
import logging
from HL7 import HL7_extraction
from selenium.common.exceptions import (
    NoSuchElementException, 
    StaleElementReferenceException
    )


def main(): 
    # Path for log file
    # log_path = os.path.expanduser('~') + '/Desktop'
    
    # initializing logging file 
    logging.basicConfig(filename='Debug_ERROR.log', level=logging.DEBUG, format='%(asctime)s:%(levelname)s:%(message)s')
    
    # asking for input to start code
    username = str(input("Please enter your TST username: "))
    password = str(input("Please enter your password: "))
    FromDate = str(input("Please enter starting date of TST export: "))
    EndDate = str(input("Please enter end date of TST export: "))
    LabName = str(input('Please enter the lab name that you want to validate: '))
    try:
        # process start 
        logging.info('Starting the process....')
    

        # logging function call
        logging.debug('Running HL7_Extraction method hl7_copy')
        report = HL7_extraction(
            username = username, #'krastegar',
            paswrd = password, #'Hamid&Mahasty1',
            FromDate= FromDate, #'02/13/2023',
            ToDate= EndDate, #'02/24/2023',
            LabName= LabName #"Palomar Health Downtown Campus Laboratory NEW"
            # LOGAN HEIGHTS FAMILY HEALTH CENTER LAB NEW
            # Palomar Health Downtown Campus Laboratory NEW
            )
        # setup = report.install()
        di = report.hl7_copy()
        break_pt = None
    except NoSuchElementException as ne:
        logging.error('An error occurred: %s', str(ne), exc_info=True)
        input("Copy Error Report and then press Enter")
    
    except StaleElementReferenceException as se:
        logging.error('An error occurred: %s', str(se), exc_info=True)
        input("Copy Error Report and then press Enter")
    
    logging.info('Process Complete...')
    '''
krastegar
Hamid&Mahasty1
02/13/2023
02/24/2023
LOGAN HEIGHTS FAMILY HEALTH CENTER LAB NEW
    '''
    return

if __name__=="__main__":
    sys.exit(main())

