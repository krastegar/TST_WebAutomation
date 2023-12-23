# ---------------------------------------------------------------------------------------------------
#   Purpose: 
#       This script will automate TST WebCMR validation with HL7 messages received from a specific lab.
#       Validation is completed by comparing sections in the HL7 message with the sections in the 
#       TSTWebCMR Demographics and Laboratory sections in the Disease Incidents page.
#   Algorithm:
#       1. Login to TST
#       2. Search for lab
#       3. Export Incoming Messages in Bulk
#       4. Classify cases into 3 categories: Categorical, Numeric, Mixed
#       5. Scrape TSTWebCMR Demographics and Laboratory sections for specific cases
#       6. Validate with HL7 messages
#       7. Produce Summary Excel Report
#       8. Logout
#   
#   Author: Kiarash Rastegar
#   Date: 2023-12-22
#   Version: 1.0
# ---------------------------------------------------------------------------------------------------


import sys

import logging
from HL7 import HL7_extraction
from selenium.common.exceptions import (
    NoSuchElementException, 
    StaleElementReferenceException,
    TimeoutException,
    SessionNotCreatedException,
    WebDriverException
    )


def main():
    # Set up logging configuration
    logging.basicConfig(filename='Log_Info.log', level=logging.INFO, format='%(asctime)s:%(levelname)s:%(message)s')

    # Prompt the user for input
    username = input("Please enter your TST username: ")
    password = input("Please enter your password: ")
    FromDate = input("Please enter starting date of TST export: ")
    EndDate = input("Please enter end date of TST export: ")
    LabName = input('Please enter the lab name that you want to validate: ')

    # Log user parameters
    logging.info(
        f'''
        ------ User Parameters ------
        USER_NAME : {username}
        PASSWORD : {password}
        START_DATE : {FromDate}
        END_DATE : {EndDate}
        LAB_NAME : {LabName}
        -----------------------------
        '''
    )

    try:
        # Log the start of the process
        logging.info('Starting the process....')

        # Call the HL7_extraction function with the user parameters
        report = HL7_extraction(
            username=username,
            paswrd=password,
            FromDate=FromDate,
            ToDate=EndDate,
            LabName=LabName
        )

        # Copy the HL7 report
        di = report.hl7_copy()

    except NoSuchElementException as ne:
        # Log the exception and prompt the user to check the log
        logging.exception("An error occurred, check Log_info.log: %s", ne)
        input('Check log info...press enter after complete')
    except StaleElementReferenceException as se:
        # Log the exception and prompt the user to check the log
        logging.exception("An error occurred, check Log_info.log: %s", se)
        input('Check log info...press enter after complete')
    except TimeoutException as te:
        # Log the exception and prompt the user to check the log
        logging.exception("An error occurred, check Log_info.log: %s", te)
        input('Check log info...press enter after complete.')
    except SessionNotCreatedException as noSession:
        # Log the exception related to session creation and compatibility
        logging.exception("Incompatibility with Chromedriver and Chromebrowser: %s", noSession)
    except WebDriverException as sessionIncompatible:
        # Log the exception related to session compatibility
        logging.exception("Incompatibility with Chromedriver and Chromebrowser: %s", sessionIncompatible)

    # Log the completion of the process
    logging.info('Process Complete...')

if __name__=="__main__":
    sys.exit(main())

