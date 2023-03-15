#-------------------------------------------------------------------------------------------
# 
#   Purpose:
#   Creating a program that will be able to take in RangeExports and IncomingMessageMonitor 
#   Exports and automate making summary reports of the export files. 
#    
#   Author: Kiarash Rastegar
#   Date: 2/27/23
#-------------------------------------------------------------------------------------------

import os
import sys 
import pandas as pd
from sklearn.metrics.pairwise import cosine_similarity
from util import (
    tstRangeQuery_demographic,
    tstRangeQuery_lab, 
    completeness_report, 
    database_connection,
    combined_MM,
    test_match
)
def main():
    # connecting python to microsoft access file
    # Need to input name of .accdb file to run pipe line
    folder_path = r'C:\Users\krastega\Desktop\MicrosoftAccessDB' # will be used as an input 
    
    # only need 1 connection for TST .accdb file 
    conn, cursor = database_connection(folder_path=folder_path,
                                       file_name='TST_DIE_02132023_03072023_exp022323')

    # Going to make one excel sheet with completeness reports from both demographic 
    # and lab information
    writer = pd.ExcelWriter('completeness_reports.xlsx', engine='xlsxwriter')

    # looking at Disease Incidents table specifically and reading it into a dataframe
    # Also making completeness report for Disease Incident Export table 
    query = tstRangeQuery_demographic(test_center_1='Palomar', test_center_2='Pomerado')
    df_demo = completeness_report(query=query, conn=conn)
    df_demo.to_excel(writer, sheet_name='Sheet1', startcol=0, index=False)
    
    # Looking at Laborartory Systems table and making completeness report
    query = tstRangeQuery_lab(test_center_1='Palomar', test_center_2='Pomerado')
    df_lab = completeness_report(query=query, conn=conn)
    df_lab.to_excel(writer, sheet_name='Sheet1', startcol = len(df_demo.columns)+1, index=False)

    # close writer object
    writer.close()

    # take incoming message monitors and produce a list of entries 
    # that were found in Production and not in TST
    prod_df = combined_MM(folder_path=folder_path,
                          startswith='PROD')
    tst_df = combined_MM(folder_path=folder_path,
                          startswith='TST')
    # matching message monitor
    # I want to match on these columns(prod_df[['DILR_ResultTest', 'DILR Accession']])
    # need to mess with TST env accession numbers to get rid of leading 0000

    # right join on Accession number to get matching from both datasets 
    merged_df = pd.merge(
         pd.DataFrame(tst_df[['DILR_AccessionNumber', 'DILR_ResultTest']]), 
         pd.DataFrame(prod_df[['DILR_AccessionNumber', 'DILR_ResultTest']]), 
         on=['DILR_AccessionNumber'], 
         how='right'
         )
    similarity_key = test_match(
         merged_df, 
         'DILR_ResultTest_x', 
         'DILR_ResultTest_y'
         )

    # transform matched resulted tests to their production names and 
    # do a 'true match' on Accession Number and ResultedTest
    # Using similarity_key dictionary to map ResultedTest 
    tst_df['DILR_ResultTest'].replace(similarity_key, inplace=True)
    tst_df.to_excel('TransformedTST.xlsx')

    # performing match on accession number and ResultedTest
    new_findings = pd.merge(
        tst_df, 
        prod_df, 
        on= ['DILR_AccessionNumber', 'DILR_ResultTest'],
        how='right')
    new_findings.to_excel('MERGER.xlsx')

    # filtering production data frame based off of matched findings
    filtered_df = prod_df[~prod_df.isin(new_findings[['DILR_AccessionNumber', 'DILR_ResultTest']])]
    filtered_df.to_excel('Final.xlsx')

    return 
    

if __name__=="__main__":
    sys.exit(main())