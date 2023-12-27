#-------------------------------------------------------------------------------------------
# 
#   Purpose:
#   Creating a program that will be able to take in RangeExports and IncomingMessageMonitor 
#   Exports and automate making summary reports of the export files. 
#   
#   Algorithm:
#       1. Find folder where .accdb file and message monitor (MM) excel files are located
#       2. Create DB connection with TST range export (.accdb extension)      
#       3. Extract data from TST range export and calculate completeness for specified fields
#       4. Produce completeness report as an excel file called "completeness_report.xlsx"
#       5. Combine all of the MM export from TST into a combined MM export, repeat this for PROD 
#          exports
#       6. Merge both combined dataframes and perform matching algorithm using test_match()
#       7. Transform TST ResultTest types to there corresponding PROD test types based on matching
#          algorithm
#       8. Remove unseen test types in PROD that were not seen in TST and perform merge on DILR_ResultTest and DILR_AccessionNumber
#       9. Subtract the ones that are seen both PROD and TST 
#
#   Author: Kiarash Rastegar
#   Date: 2/27/23
#-------------------------------------------------------------------------------------------

import sys 

from mismatch import bulkval
from menu import Menu

def main():

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




if __name__=="__main__":
    sys.exit(main())