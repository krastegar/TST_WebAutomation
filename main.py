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