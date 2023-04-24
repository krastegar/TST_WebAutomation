import sys 

from Completeness import Completeness

def main():

    # calling Completeness object 
    report_maker = Completeness(
        file_name= 'TST_04212023_EISBDisease',
        lab_name='Point Loma Nazarene University Wellness Center',
        folder_path= 'S:\PHS\EPI\EPIRESTRICTED\BEACON\EPI_BEACON_ELR\Point Loma Nazarene University Wellness',
        test_center_1='Point Loma',
        test_center_2='Wellness Medical Center')
    report_maker.completeness_report()
    return

if __name__=="__main__":
    sys.exit(main())