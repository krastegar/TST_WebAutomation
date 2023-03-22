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
    folder_path = r'C:\Users\krastega\OneDrive - County of San Diego\Desktop\MicrosoftAcessDB' # will be used as an input 
    
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
    tst_df.to_excel('combined_tst.xlsx'); prod_df.to_excel('combined_prod.xlsx')
    # matching message monitor
    # I want to match on these columns(prod_df[['DILR_ResultTest', 'DILR Accession']])
    # need to mess with TST env accession numbers to get rid of leading 0000
    #prod_df.to_excel('combined_prod.xlsx')
    # right join on Accession number to get matching from both datasets 
    merged_df = pd.merge(
         pd.DataFrame(tst_df[['DILR_AccessionNumber', 'DILR_ResultTest']]), 
         pd.DataFrame(prod_df[['DILR_AccessionNumber', 'DILR_ResultTest']]), 
         on=['DILR_AccessionNumber'], 
         how='right'
         )
    similarity_key, missingProdTests = test_match( # used to be similarity key, but now is match df
         merged_df, 
         'DILR_ResultTest_x', 
         'DILR_ResultTest_y'
         ) 
    # taking out tests that are not seen in TST but seen in PROD
    unseen_tests = prod_df[prod_df['DILR_ResultTest'].isin(missingProdTests)] 
    #unseen_tests.to_excel('New_Prod.xlsx')

    # Transforming loinc code in tst to match prod environment 
    tst_df['DILR_ResultTest'].replace(similarity_key, inplace=True)
    #print(tst_df['DILR_ResultTest'])
    # performing match on accession number and ResultedTest
    new_findings = pd.merge(
        tst_df, 
        prod_df, 
        on= ['DILR_AccessionNumber', 'DILR_ResultTest'],
        how='left')

    # These are the matches from prod_df to tst_df after changing the tst_df ResultTest 
    # to match prod_df ResultTest
    # issues with this is that we have ResultTests seen in TST environment that are not seen Prod

    # going to filter out new_findings to get rid of tst not seen in prod
    # test_matches = list(similarity_key.values())
    # new_findings = new_findings[new_findings['DILR_ResultTest'].isin(test_matches)] 
    # now I have finally found matches for ResultTest in TST and Prod...
    # ...need to filter out these tests 

    # going to filter out non matches of accession number on new findings 
    missing_tests = prod_df[~prod_df['DILR_AccessionNumber'].isin(new_findings['DILR_AccessionNumber'])]
    final_missing_report = pd.concat([missing_tests, unseen_tests], ignore_index=True)
    final_missing_report.to_excel('final_missing_report.xlsx')

    # comparing my final_report to marjorie's
    mr_df = pd.read_excel('Palomar_Missing_Results_07MAR2023.xlsx')
    compare_df = mr_df[~mr_df['PROD_AccessionNumber'].isin(final_missing_report['DILR_AccessionNumber'])]
    compare_df.to_excel('compare.xlsx')

    return 
    

if __name__=="__main__":
    sys.exit(main())