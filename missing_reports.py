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
    def __init__(self, folder_path, lab_name):
        self.folder_path = folder_path
        self.labname = lab_name


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
    
    def frequency_table(self, column1 = 'DILR_ResultTest_x', column2 = 'DILR_ResultTest_y'):
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
    
    def test_match(self, column1 = 'DILR_ResultTest_x', column2 = 'DILR_ResultTest_y'):

        '''
        new_df structure:

                                                    SarsCov2        HepC SurfAnti A
        SarsCoronaVirusLike::PROBE:ANT                  15               0

        HepatitusC Surface Antingen::MOLC:ANTH          0                42

        Code below is looping through each row and finding the column with the highest counts, if there
        are 2 columns that have the same amount of counts it looks at both columns to see if there are 
        any other rows within that column that have higher counts. If there are then it removes the column
        name from the list of tie breaker columns. 

        Finally for the test types in production that are not seen in TST are put together in one list
        and reported as missing. 
        '''
        new_df, missing_test = self.frequency_table(column1, column2) # new_df.to_excel('Intermediate_match.xlsx')
        # new_df.to_excel('Intermediate_match.xlsx') only used for diagnostics 
        matched_pairs = {}
        match_array = []
        for index, row in new_df.iterrows():
            max_val = row.max()
            col_names = row[row == max_val].index.tolist()
            if len(col_names) > 1: 
                for col in col_names: 
                    tie_break_score = new_df[col].max()
                    if tie_break_score > max_val:  # this was originally if tie_break_score > max_val:
                        col_names.remove(new_df[col].name)
            matched_pairs[col_names[0]] = index
            match_array.append((index, col_names[0]))
        match_df = pd.DataFrame(match_array, columns=['ProdTest', 'TSTTest'])
        #match_df.to_excel('match.xlsx')
        
        # filter data tables for repeat entries in TSTTest column 
        corrected_matches_df, prod_test_mapping , mismatch_test= self.remove_duplicates(match_df)
        
        # List of missing tests (seen in Prod env but not TST) 
        # I added the cases of mismatched preliminary vs final vs correction to the missing test list
        # using list comprehension to combine both missing tests and mismatch test 
        _ = [missing_test.append(test) for test in mismatch_test]

        # I added an indicator showing why that there are some test are marked as 'missing' even though
        # a few of the accession numbers were seen in both PROD and TST message monitors 
        # I also added an indicator to signal that mismatched preliminary vs final vs correction are in 
        # Missing test lists for production 
        for index, row in corrected_matches_df.iterrows():
            prod_testName = row['ProdTest']
            tst_testName = row['TSTTest']
            if tst_testName == 'N/A' and prod_testName not in missing_test: 
                corrected_matches_df.loc[
                    index, 'TST_exceptions'
                    ] = 'Few accession # found, but no good name match'
            elif prod_testName in mismatch_test:
                corrected_matches_df.loc[
                    index, 'TST_exceptions'
                    ] = '''
                    Mismatched preliminary vs final vs correction to the missing test list.
                    These tests are put into production missing test list 
                    '''
            else: pass 
        corrected_matches_df.drop('TSTTest', axis=1, inplace = True)
        #corrected_matches_df.to_excel('OneToOne_Match.xlsx')

        return prod_test_mapping, missing_test, corrected_matches_df
    
    def remove_duplicates(self, df):
        '''
        After doing the initial match based on frequency, we are going to look at the 
        tests that have duplicate matches to PROD environment and filter out the best 
        matches 
        '''
        final_pattern = r'^(?=.*?- Final\b).+'
        correction_results = r'^(?=.*?- Correction to results\b).+'
        preliminary = r'^(?=.*?- Preliminary\b).+'
        
        mismatch_test = []
        # checking to see if there are any 'prelimanary' or 'Correction to results'
        for index, row in df.iterrows(): 
            tst, prod = str(row['TSTTest']), str(row['ProdTest'])
            if re.match(final_pattern, prod) and re.match(final_pattern,tst): 
                df.loc[index, 'TST'] = tst
                #print('MATCHHH')
            elif re.match(correction_results, prod) and re.match(correction_results,tst):
                df.loc[index, 'TST'] = tst
            elif re.match(preliminary, prod) and re.match(preliminary,tst):
                df.loc[index, 'TST'] = tst
            elif all([not re.match(preliminary, prod), re.match(preliminary, tst)]):
                mismatch_test.append(prod) #continue
            elif all([not re.match(correction_results, prod), re.match(correction_results,tst)]):
                mismatch_test.append(prod) #continue
            elif all([not re.match(final_pattern, prod), re.match(final_pattern,tst)]):
                mismatch_test.append(prod) #continue
                # this condition says if one suffix is different than the other  than
            else: 
                #mismatch_test.append(prod)
                df.loc[index, 'TST'] = np.nan

        # finally make a key:value dictionary with Production ResultTest and TST ResultTest
        prod_tst_mapping = {row['TST']: row['ProdTest'] for index, row in df.iterrows()}

        return df, prod_tst_mapping, mismatch_test
    
    
    def missing_test_report(self, prodTestsMissingInTST):
        '''
        Function is made to make a missing test report from both tests seen in TST
        and not in PROD and visa versa. 
        Parameter (tst_df): This is the TST combinded dataframe that already has the DILR_ResultTest
                            column already transformed after matching algorithm
        Parameter (prod_df): This is just the combined dataframe from all of the prod message monitors
        Parameter (prodTestMissingInTST): This is a list that is produced by test_match() that has
                                        test types that are seen in PROD but not seen in TST. Produced 
                                        after matching algorithm.
        '''
        # gettting combined tst and prod dataframes
        tst_df, prod_df = self.prod_tst_combinedDF()

        # Unique Test Results seen in both TST and PROD 
        tst_unique = list(np.unique(tst_df['DILR_ResultTest']))
        prod_unique = list(np.unique(prod_df['DILR_ResultTest']))

        tst_notProd = []
        for test in tst_unique:
            if test not in prod_unique:
                tst_notProd.append(test)
                
        # building dataframe for both missing test lists 
        df_1 = pd.DataFrame({'PROD not in TST': prodTestsMissingInTST})
        df_2 = pd.DataFrame({'TST not in PROD': tst_notProd})
        missing_report_df = pd.concat([df_1, df_2], axis=1)
        #missing_report_df.to_excel('missing_test_types.xlsx')

        return missing_report_df
    

    def final_missing(self):
        '''
        Method which generates a list of test missing in TST but seen in PROD (prodTestsMissingInTST), 
        as well as a summary dataframe of how the test names were matched in each environment (one2one_df).
        It transforms the tests that were matched from PROD to TST into the names that they should have in
        PROD and then keeps any of the test with accession numbers, only, that are not seen in the new matched
        dataframe.
        '''
        similarity_key, prodTestsMissingInTST, one2one_df = self.test_match( 
            column1='DILR_ResultTest_x', 
            column2='DILR_ResultTest_y'
        )
        tst_df, prod_df = self.prod_tst_combinedDF()
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
        
        # going to filter out non matches of accession number on new findings 
        final_missing_report = prod_df[~prod_df['DILR_AccessionNumber'].isin(new_findings['DILR_AccessionNumber'])]
        
        return final_missing_report, prodTestsMissingInTST, one2one_df

    def prod_tst_combinedDF(self):
        tst_df = self.combined_MM(start_with='TST')
        prod_df = self.combined_MM(start_with='PROD')
        return tst_df,prod_df
    
    def sanity_check(self, final_missing_report):
        '''
        Function looks at the final missingn reports list and checks if any of the accession numbers in 
        that list that are seen in combined TST message monitor (MM) exports. If there are it produces an 
        excel that is full of these lists and sees what the tests are labeled as in Combined Production MM
        
        Note:
            What was found in the tst leaks variable were a bunch of mislabeled tests in TST MM that 
            remove_duplicates() took out from the final_missing report, but since these accession numbers
            map back to TST and are technically 'Reported' I counted them as missing since they are mislabeled.
            A summary of these mislabeling is shown in OneToOne_Match.xlsx file that is produced when running
            the script
        '''

        # getting combined TST and PROD dataframes
        tst_df, prod_df = self.prod_tst_combinedDF()

        # Variable to find "leaks" in my match and find algorithm 
        # i.e) tests with accession numbers in TST that I said were missing
        tst_leaks =tst_df[tst_df['DILR_AccessionNumber'].isin(final_missing_report['DILR_AccessionNumber'])]

        # removing the exceptions from the final missing report 
        final_missing_report = final_missing_report[
            ~final_missing_report['DILR_AccessionNumber'].isin(tst_leaks['DILR_AccessionNumber'])
        ]

        # Noticing that my code has some duplicates that are not accurate..
        # Going to see if dropping duplicates helps 

        # Seeing what these missing tests are in PROD environment
        prod_replacements = prod_df[prod_df['DILR_AccessionNumber'].isin(tst_leaks['DILR_AccessionNumber'])]
        #prod_replacements.to_excel("Prod_doubleCheck.xlsx")

        # dropping duplicates from final results.
        final_missing_report.drop_duplicates(['DILR_AccessionNumber', 'DILR_ResultTest'], inplace=True, keep='first')
        #final_missing_report.to_excel('final_missing_report.xlsx')

        return final_missing_report
    
    def report_builder(self):
        '''
        Method which puts all of the summary dataframes into one excel sheet
        '''
        # getting list of initial missing reports, tests seen in PROD but not TST
        # and a summary of tests that were matched in TST but not prod 
        missing_report, prod_missing_test ,match_summary = self.final_missing()

        # getting missing test types report
        missing_test_types = self.missing_test_report(prod_missing_test)

        # removing exceptions from missing report lists 
        final_missing_report = self.sanity_check(missing_report)

        # building excel sheet with all reports
        report_name = re.sub(r'[^\w\s]+', '_',self.labname)
        writer = pd.ExcelWriter(f'{report_name}_validation_reports.xlsx', engine='xlsxwriter')

        # list of missing test types
        missing_test_types.to_excel(
            writer, 
            sheet_name='Missing_TestTypes', 
            startcol=0, 
            index=False
            )
        
        # Matching summary
        match_summary.to_excel(
            writer, 
            sheet_name='One2One_matches', 
            startcol=0, 
            index=False
            )
        
        # final missing reports list
        final_missing_report.to_excel(
            writer, 
            sheet_name='Tests in Prod not TST', 
            startcol=0, 
            index=False
            )
        writer.close()

        return 