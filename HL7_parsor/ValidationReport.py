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
    logging.basicConfig(filename='ERROR.log', level=logging.ERROR)
    
    # asking for input to start code

    try:
        username = str(input("Please enter your TST username: "))
        password = str(input("Please enter your password: "))
        FromDate = str(input("Please enter starting date of TST export: "))
        EndDate = str(input("Please enter end date of TST export: "))
        LabName = str(input('Please enter the lab name that you want to validate: '))
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
    except NoSuchElementException as e:
        logging.error('An error occurred: %s', str(e))
        input("Copy Error Report and then press Enter")
    
    except StaleElementReferenceException as error:
        logging.error('An error occurred: %s', str(error))
        input("Copy Error Report and then press Enter")

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

