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
import pandas as pd

from util import (
    tstRangeQuery_demographic,
    tstRangeQuery_lab, 
    completeness_report, 
    database_connection,
    combined_MM,
    test_match,
    sanity_check,
    missing_test_report
)

def main():
    
    # connecting python to microsoft access file
    # Need to input name of .accdb file to run pipe line
    folder_path = r'C:\Users\krastega\OneDrive - County of San Diego\Desktop\MicrosoftAcessDB' # will be used as an input 
    
    # only need 1 connection for TST .accdb file 
    conn, _ = database_connection(folder_path=folder_path,
                                       file_name='TST_DIE_04202023_05042023')

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
    prod_df.reset_index(inplace=True); tst_df.reset_index(inplace=True)
    prod_df.to_excel('combined_prod.xlsx'); tst_df.to_excel('combined_tst.xlsx')
    # right join on Accession number to get matching from both datasets 
    merged_df = pd.merge(
         pd.DataFrame(tst_df[['DILR_AccessionNumber', 'DILR_ResultTest']]), 
         pd.DataFrame(prod_df[['DILR_AccessionNumber', 'DILR_ResultTest']]), 
         on=['DILR_AccessionNumber'], 
         how='right'
         )
    similarity_key, prodTestsMissingInTST, one2one_df = test_match( # used to be similarity key, but now is match df
         merged_df, 
         'DILR_ResultTest_x', 
         'DILR_ResultTest_y'
         )

    # Transforming loinc code in tst to match prod environment 
    tst_df['DILR_ResultTest'].replace(similarity_key, inplace=True)

    # performing match on accession number and ResultedTest
    # These are the matches from tst_df to prod_df after changing the tst_df ResultTest 
    new_findings = pd.merge(
        tst_df, 
        prod_df, 
        on= ['DILR_AccessionNumber', 'DILR_ResultTest'],
        how='left', # change this back to left to match previous results
        indicator=True)
    new_findings.reset_index(inplace=True)

    # Creating missing test report from test seen in TST env but not seen in PROD
    # Also test seen in PROD but not seen in TST 
    _ = missing_test_report(tst_df,prod_df, prodTestsMissingInTST)

    # going to filter out non matches of accession number on new findings 
    final_missing_report = prod_df[~prod_df['DILR_AccessionNumber'].isin(new_findings['DILR_AccessionNumber'])]
    
    # See how this compares with combined tst dataframes  
    result_check = sanity_check(prod_df, tst_df, final_missing_report)


    # Cross Comparison with Marjorie missing list 


if __name__=="__main__":
    sys.exit(main())