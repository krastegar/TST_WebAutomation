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

from missing_reports import BulkVal
from menu import Menu

def main():

    # calling my menu object
    menu = Menu()

    # getting user input
    folder_path = menu.get_input("Enter folder_path: ", str, lambda x: menu.is_valid_folder_path(x))
    lab_name = menu.get_input('Name of Lab: ', str)

    # calling bulkvalidation object 
    report_maker = BulkVal(folder_path=folder_path, lab_name=lab_name)
    # C:\Users\krastega\OneDrive - County of San Diego\Desktop\MicrosoftAcessDB

    # building report of all summary dataframes
    report_maker.report_builder()

    return

if __name__=="__main__":
    sys.exit(main())