#-------------------------------------------------------------------------------------------
# 
#   Purpose:
#   Creating a program that will be able to take in Incoming Message Monitor 
#   Exports and automate making summary reports of the export files. 
#   
#   Algorithm:
#       1. Find folder where message monitor (MM) excel files are located
#       2. Combine all of the MM export from TST into a combined MM export, repeat this for PROD 
#          exports
#       3. Merge both combined dataframes and perform matching algorithm using test_match()
#       4. Transform TST ResultTest types to there corresponding PROD test types based on matching
#          algorithm
#       5. Remove unseen test types in PROD that were not seen in TST and perform merge on DILR_ResultTest and DILR_AccessionNumber
#       6. Subtract the ones that are seen both PROD and TST 
#
#   Author: Kiarash Rastegar
#   Date: 2/27/23
#-------------------------------------------------------------------------------------------

import sys 
import logging
from mismatch import bulkval
from menu import Menu

def main():

    # Logging process
    logging.basicConfig(filename='ErrorLog.log', level=logging.ERROR, format='%(asctime)s:%(levelname)s:%(message)s')

    try: 
        # calling my menu object
        menu = Menu()

        # getting user input
        folder_path = menu.get_input("Enter folder_path: ", str, lambda x: menu.is_valid_folder_path(x))
        lab_name = menu.get_input('Name you want the report to be: ', str)
        
        # C:\Users\krastega\OneDrive - County of San Diego\Desktop\MicrosoftAcessDB
        validator = bulkval(folder_path=folder_path, lab_name=lab_name)
        
        # take incoming message monitors and produce a list of entries 
        # that were found in Production and not in TST
        prod_df = validator.combined_MM(startswith='PROD')
        tst_df = validator.combined_MM(startswith='TST')
        prod_df.reset_index(inplace=True); tst_df.reset_index(inplace=True)
        
        # producing final report of missing test types, missing reports, and summary of matching algorithm
        validator.report_builder(prod_df, tst_df)
    
    except: # catch any errors and log them along with the traceback
        logging.error(sys.exc_info())



if __name__=="__main__":
    sys.exit(main())