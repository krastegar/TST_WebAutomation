import sys

from DI_Incident import DI_Search
from HL7 import HL7_extraction


def main(): 

    # Login and go to staging area
    report = HL7_extraction(
        username = 'krastegar',
        paswrd = 'Hamid&Mahasty1',
        FromDate='01/18/2023',
        ToDate='01/18/2023',
        LabName="LOGAN HEIGHTS FAMILY HEALTH CENTER LAB NEW"
        )
    di = report.hl7_copy()
    
    '''
    Some issues maybe in the future: 
    1.) if the website updates the code will break, because I am finding these elements with there ID
    2.) The code runs continuously without stopping. 
        - Issue with this is that if there are a lot of unique tests then there will be a lot of searches
        - Currently I am doing unique combinations of ResultedTest, ResultValue, and my own Classification column
    3.) If none of the incoming messages are imported we will get empty excel files or the code will break
    4.) might need to refresh downloads folder before running code (not sure)
    '''
    return

if __name__=="__main__":
    sys.exit(main())

