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
import numpy as np
import gensim
from sklearn.metrics.pairwise import cosine_similarity
from util import (
    tstRangeQuery_demographic,
    tstRangeQuery_lab, 
    completeness_report, 
    database_connection,
    combined_MM
)
def main():
    # connecting python to microsoft access file
    # Need to input name of .accdb file to run pipe line
    folder_path = r'C:\Users\krastega\Desktop\MicrosoftAccessDB' # will be used as an input 
    
    # only need 1 connection for TST .accdb file 
    conn, cursor = database_connection(folder_path=folder_path,
                                       file_name='TST_DIE_02132023_02212023_exp022323')

    # looking at Disease Incidents table specifically and reading it into a dataframe
    # Also making completeness report for Disease Incident Export table 
    query = tstRangeQuery_demographic(test_center_1='Palomar', test_center_2='Pomerado')
    df_demo = completeness_report(query=query, conn=conn)

    # Looking at Laborartory Systems table and making completeness report
    query = tstRangeQuery_lab(test_center_1='Palomar', test_center_2='Pomerado')
    df_lab = completeness_report(query=query, conn=conn)

    # take incoming message monitors and produce a list of entries 
    # that were found in Production and not in TST
    prod_df = combined_MM(folder_path=folder_path,
                          startswith='PROD')
    tst_df = combined_MM(folder_path=folder_path,
                          startswith='TST')
    # matching message monitor
    # I want to match on these columns(prod_df[['DILR_ResultTest', 'DILR Accession']])
    # need to mess with TST env accession numbers to get rid of leading 0000

    # inner join on Accession number to get matching from both datasets 
    merged_df = pd.merge(
         pd.DataFrame(tst_df[['DILR_AccessionNumber', 'DILR_ResultedTest']]), 
         pd.DataFrame(prod_df[['DILR_AccessionNumber', 'DILR_ResultedTest']]), 
         on=['DILR_AccessionNumber'], 
         how='inner'
         )
    similarity_df = word2vec(
         merged_df, 
         'DILR_ResultedTest_x', 
         'DILR_ResultedTest_y'
         )
    similarity_df.to_excel('matchv1.xlsx')
    return 
    
def word2vec(df, column1, column2):
    counts = {}
    # creating count table of unique Prodtest and TSTtest that are seen together
    # And I am going to use these counts to find matches of test names 
    for i, row in df.iterrows():
        col1_val = row[column1]
        col2_val = row[column2]
        
        if col1_val not in counts:
            counts[col1_val] = {}
        
        if col2_val in counts[col1_val]:
            counts[col1_val][col2_val] += 1
        else:
            counts[col1_val][col2_val] = 1
    # Now going to look at the the dictionary values and look inside the nested dictionary 
    # for counts and which ever counts is the most for that value pair I will use that as a match
    new_df = pd.DataFrame(counts)
    matched_array = []
    for index, row in new_df.iterrows():
        max_val = row.max()
        col_names = row[row == max_val].index.tolist()
        if len(col_names) > 1: 
            row_names = index
            for col in col_names: 
                tie_break_score = new_df[col].max()
                if tie_break_score > max_val: 
                    col_names.remove(new_df[col].name)
        matched_array.append((index, col_names[0], max_val))
    
    matched_df = pd.DataFrame(matched_array, columns=['PRODTest', 'TSTTest', 'Frequency of Match'])
        
    #print(counts)
    return matched_df

if __name__=="__main__":
    sys.exit(main())