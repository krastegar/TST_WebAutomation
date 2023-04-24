#-------------------------------------------------------------------------------------------
# 
#   Purpose:
#   Looking at both Message Monitor (MM) exports, both as in PROD and TST environments. We look to see 
#   which Tests are seen in PROD environment but not seen in TST and vice versa. The class produces 
#   a a list of final missing report of tests seen in PROD but not in TST, missing test types not seen in PROD
#   but seen in TST and vice versa, and a summary of how the program matches test from one environment to another.
#   There are some exceptions when it comes to matching test, which are summarized in the One2One_match.xlsx
#   
#   Algorithm:
#       1. Find folder where all the MM exports are seen
#       2. Combine all of the TST MM exports into one file and do the same for PROD MM exports as well
#       3. Join combined exports to each other only using accession number, only joining accession numbers
#          that match accession numbers in PROD (right join) 
#       4. Look at the most frequently matched pairs of accession number and result tests to determine which
#          test names from TST environment belong to PROD environment 
#       5. Create a missing test type list based on PROD tests that do not have a match in TST environment, which
#          also includes test that do not have the same ending tag i.e) Final vs Prelimenary vs Correction
#       6. Double check if the tests not seen in PROD are truly missing or if they have more missing than matches
#          and mark these test types as exceptions in One2One_match summary report
#       7. Replace all of the matched TST names to their corresponding PROD test names in TST environment
#       8. Merge transformed TST exports with PROD on both accession number and test name
#       9. Filter out PROD combined exports by the accession numbers that do not matched the newly merged 
#          dataframe
#       10. Run another check to make sure that there were not any tests that my program says are missing and 
#           are actually seen in TST environments
#       11. Takes the missing list of test types seen in PROD but not in TST and uses it to make a list of test 
#           types seen in TST but not in PROD. 
#       12. Combines both lists into a dataframe and produces a missing_testtypes report in Excel. 
#
#   Author: Kiarash Rastegar
#   Date: 4/20/23
#-------------------------------------------------------------------------------------------

import pandas as pd
import numpy as np
import os
import re

class BulkVal:
    def __init__(self, folder_path):
        self.folder_path = folder_path


    def combined_MM(self, start_with):
        '''
        Function that looks into a directory and looks for excel files that start with 'startwith' 
        variable. Grabs all of the file names and reads them into pandas df, which is later 
        concatenated 
        '''
        # getting list of files for either TST or Prod IMM exports
        prod_file_list = [
            os.path.join(self.folder_path, file) for file in os.listdir(self.folder_path)
            if os.path.basename(file).startswith(start_with) and os.path.basename(file).endswith('.xlsx')
            ]

        # getting column names from one of the data frames in a specific environment 
        test_df = pd.read_excel(prod_file_list[0])
        combined_df = pd.DataFrame(columns=test_df.columns)
        
        # combining all of the excel files into 1 df
        prod_db_array = []

        for file in prod_file_list: 
            df = pd.read_excel(file)
            prod_db_array.append(df)
        combined_df = pd.concat(prod_db_array)

        return combined_df
    
    def merge2match(self):
        '''
        After combining all of the exports from both environments into there own combined exports, 
        we merged the combine exports which we will then use for the matching algorithm test match
        '''

        # Creating combined dataframes
        combined_tst = self.combined_MM(start_with='TST')
        combined_prod = self.combined_MM(start_with='PROD')

        # Merging combined exports
        merged_df = pd.merge(
         pd.DataFrame(combined_tst[['DILR_AccessionNumber', 'DILR_ResultTest']]), 
         pd.DataFrame(combined_prod[['DILR_AccessionNumber', 'DILR_ResultTest']]), 
         on=['DILR_AccessionNumber'], 
         how='right'
         )

        return merged_df
    
    def frequency_table(self, merged_df):
        '''
        A function that loops through a Merged dataframe. This merged dataframe is merged on the 
        accession number column, which results in having two ResultedTest columns. The function loops 
        through these columns and sees how frequently each column pairs are seen next to each other
        i.e)
            ProdTest            TSTTest
        SARSCov2            SARS Coronavirus+Like SARS
        SARSCov2            Influenza Virus A 
        ...                 ...
        It takes the most frequently seen unique pairs and makes a dictionary that will be used as
        a key to filter out tests that do not match 
        '''
        df = self.merge2match()
        counts = {}
        # creating count table of unique Prodtest and TSTtest that are seen together
        # And I am going to use these counts to find matches of test names 
        for _, row in df.iterrows():
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

        # going to look at the empty column of new_df and use those index's 
        # that have only empty entries as missing and put those as missing test types

        new_df.columns = new_df.columns.fillna('N/A') # changing empty column name to N/A
        all_other_cols = [col for col in new_df.columns if col != 'N/A']
        missing_test = []
        for index, row in new_df.iterrows():
            if row['N/A'] is not np.nan and all(row[all_other_cols].isna()):
                missing_test.append(index)
                
        return new_df, missing_test