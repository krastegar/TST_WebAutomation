#-------------------------------------------------------------------------------------------
# 
#   Purpose:
#   Create an external python file that contains functions that will be useful for analyzing
#   validation data from Range Export and Incoming Message Monitor Export
#    
#   Author: Kiarash Rastegar
#   Date: 2/27/23
#-------------------------------------------------------------------------------------------

import pandas as pd
import numpy as np
import os
import re
import xlsxwriter

class bulkval:
    def __init__(self, folder_path, lab_name) -> xlsxwriter:
        self.folder_path = folder_path
        self.lab_name = lab_name

    def combined_MM(self, startswith):
        """
        Generate a combined dataframe from multiple Excel files.

        Parameters:
            startswith (str): The prefix of the file names to include in the combined dataframe.

        Returns:
            combined_df (pandas.DataFrame): The combined dataframe containing data from all the Excel files.

        Raises:
            AssertionError: If there are no files that start with the specified prefix.
        """

        # getting list of files for either TST or Prod IMM exports
        prod_file_list = [os.path.join(self.folder_path, file) for file in os.listdir(self.folder_path)
                        if os.path.basename(file).startswith(startswith) and os.path.basename(file).endswith('.xlsx')]

        assert len(prod_file_list) > 0, f'There were not any files that start with {startswith}'

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

    def test_match(self, df, column1, column2):

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

        Parameters:
            df (DataFrame): The input DataFrame.
            column1 (str): The name of the first column.
            column2 (str): The name of the second column.

        Returns:
            prod_test_mapping (dict): A dictionary of matched pairs.
            missing_test (list): A list of missing tests.
            corrected_matches_df (DataFrame): A DataFrame of corrected matches.

        Step-by-step explanation:
        1. Generate a frequency table for the given DataFrame `df` using the columns `column1` and `column2`.
        2. Store the frequency table in the variable `new_df`.
        3. Initialize an empty dictionary `matched_pairs` to store the matched pairs.
        4. Initialize an empty list `match_array` to store the matched pairs.
        5. Iterate over each row in the frequency table `new_df`.
            - Find the maximum value in the row and the corresponding column names with that maximum value.
            - If there are multiple column names with the maximum value, perform tie-breaker logic.
            - Add the matched pair to the `matched_pairs` dictionary and the `match_array` list.
        6. Convert the `match_array` list to a DataFrame named `match_df`.
        7. Call the `remove_duplicates` function to filter the data tables for repeat entries in the TSTTest column.
        8. Append the values from `mismatch_test` to the `missing_test` list.
        9. Iterate over the rows of the `corrected_matches_df` DataFrame.
            - Check if the `tst_testName` is 'N/A' and if the `prod_testName` is not in the `missing_test` list.
            - If the conditions are met, update the 'TST_exceptions' column of the `corrected_matches_df` DataFrame.
            - If the `prod_testName` is in the `mismatch_test` list, update the 'TST_exceptions' column with a specific message.
        10. Drop the 'TSTTest' column from the `corrected_matches_df` DataFrame.
        11. Return the `prod_test_mapping`, `missing_test`, and `corrected_matches_df` as the output.

        '''
        new_df, missing_test = self.frequency_table(df, column1, column2) # new_df.to_excel('Intermediate_match.xlsx')
        #new_df.to_excel('Intermediate_match.xlsx')
        matched_pairs = {}
        match_array = []
        for index, row in new_df.iterrows():
            max_val = row.max()
            col_names = row[row == max_val].index.tolist() # picking the column name with the max value
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
        corrected_matches_df, prod_test_mapping , mismatch_test= self.remove_duplicates(match_df, missing_test)
        
        # List of missing tests (seen in Prod env but not TST) 
        # I added the cases of mismatched preliminary vs final vs correction to the missing test list
        for test in mismatch_test:
            missing_test.append(test)

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

    def remove_duplicates(self, df, missing_test):
        """
        Remove duplicates from the dataframe and create a mapping of Production ResultTest and 
        TST ResultTest. After doing the initial match based on frequency, we are going to look at the 
        tests that have duplicate matches to PROD environment and filter out the best matches 

        Parameters:
        - df: The dataframe to remove duplicates from.

        Returns:
        - df: The dataframe after removing duplicates.
        - prod_tst_mapping: A dictionary mapping Production ResultTest to TST ResultTest.
        - mismatch_test: A list of mismatched test results with the endings 'Final', 'Correction to results', and 'Preliminary'.
        """

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

    def frequency_table(self, df, column1, column2):
        '''
        A function that loops through a Merged dataframe. This merged dataframe is merged on the 
        accession number column, which results in having two ResultedTest columns. The function loops 
        through these columns and sees how frequently each column pairs are seen next to each other
        i.e)
            ProdTest            TSTTest
        SARSCov2            SARS Coronavirus+Like SARS
        SARSCov2            Influenza Virus A 
                            ...
        It takes the most frequently seen unique pairs and makes a dictionary that will be used as
        a key to filter out tests that do not match 
        '''
        
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


    def sanity_check(self, prod_df, tst_df, final_missing_report):
        '''
        Function looks at the final missing reports list and checks if any of the accession numbers in 
        that list that are seen in combined TST message monitor (MM) exports. If there are it produces an 
        excel that is full of these lists and sees what the tests are labeled as in Combined Production MM
        
        Note:
            What was found in the tst leaks variable were a bunch of mislabeled tests in TST MM that 
            remove_duplicates() took out from the final_missing report, but since these accession numbers
            map back to TST and are technically 'Reported' I counted them as missing since they are mislabeled.
            A summary of these mislabeling is shown in OneToOne_Match.xlsx file that is produced when running
            the script
        '''
        # Variable to find "leaks" in my match and find algorithm 
        # i.e) tests with accession numbers in 
        tst_leaks =tst_df[tst_df['DILR_AccessionNumber'].isin(final_missing_report['DILR_AccessionNumber'])]

        # removing the exceptions from the final missing report 
        final_missing_report = final_missing_report[
            ~final_missing_report['DILR_AccessionNumber'].isin(tst_leaks['DILR_AccessionNumber'])
        ]

        # Noticing that my code has some duplicates that are not accurate.

        # dropping duplicates from final results.
        final_missing_report.drop_duplicates(['DILR_AccessionNumber', 'DILR_ResultTest'], inplace=True, keep='first')
        

        return final_missing_report

    def missing_test_report(self, tst_df, prod_df, prodTestsMissingInTST):
        '''
        Function is made to make a missing test report from both tests seen in TST
        and not in PROD and visa versa. 
        
        Args: 
            - tst_df: This is the TST combinded dataframe that already has the DILR_ResultTest
                            column already transformed after matching algorithm
            - prod_df: This is just the combined dataframe from all of the prod message monitors
            - prodTestMissingInTST: This is a list that is produced by test_match() that has
                                        test types that are seen in PROD but not seen in TST. Produced 
                                        after matching algorithm.
        Returns:
            - missing_report_df: This is a dataframe that has all the missing test types from 
                                 both TST and PROD
        '''
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
    
    def final_summary_reports(self, prod_df, tst_df):
        """
        Generate the final summary reports.

        This function takes in two dataframes, `prod_df` and `tst_df`, representing data from the production environment
        and the test environment, respectively. It performs the following operations to generate the final summary reports:

        1. Merge the `tst_df` and `prod_df` dataframes on the 'DILR_AccessionNumber' column, creating a new merged dataframe.
        This allows finding all the reports that have the same accession number in both environments.

        2. Stack the tests together in the merge dataframe to identify the test names that are most frequently seen together.
        These test names will be used as key-value pairs for transforming the 'DILR_ResultTest' in `tst_df` to match
        the 'DILR_ResultTest' in the production environment.

        3. Replace the 'DILR_ResultTest' in `tst_df` with the transformed values using the similarity key.

        4. Perform a left join on the 'DILR_AccessionNumber' and 'DILR_ResultTest' columns between `tst_df` and `prod_df`.
        This creates a new dataframe, `new_findings`, which contains the matches between the two dataframes.

        5. Reset the index of the `new_findings` dataframe.

        6. Create a missing test report by identifying the tests seen in the test environment (`tst_df`) but not seen in the
        production environment (`prod_df`), as well as tests seen in the production environment but not seen in the test
        environment.

        7. Filter out the non-matching accession numbers from `prod_df` using the `new_findings` dataframe.

        8. Perform a sanity check to identify if there are any 'ResultTests' marked as missing that are actually present in
        the test environment (`tst_df`).

        Returns:
            tuple: A tuple containing the following:
                - one2one_df (pandas.DataFrame): The dataframe with the matched test names from `tst_df` and `prod_df`.
                - final_missing (pandas.DataFrame): The dataframe containing the missing test reports.
                - missing_test_types (pandas.DataFrame): The dataframe containing the missing test types.
        """
        # need to create a merged dataframe from both tst and prod environments so that I can find all of
        # the reports that have the same accession number in both environments 
        merged_df = pd.merge(
            pd.DataFrame(tst_df[['DILR_AccessionNumber', 'DILR_ResultTest']]), 
            pd.DataFrame(prod_df[['DILR_AccessionNumber', 'DILR_ResultTest']]), 
            on=['DILR_AccessionNumber'], 
            how='right'
            )
        
        # stacking the tests together in the merge dataframe, so we can see which test names are 
        # most frequently seen together. These tests are going to be key-value pairs which will 
        # be used for transformation of TST ResultTest into PROD ResultTest 
        similarity_key, prodTestsMissingInTST, one2one_df = self.test_match( # used to be similarity key, but now is match df
            merged_df, 
            'DILR_ResultTest_x', 
            'DILR_ResultTest_y'
            )

        # Transforming ResultTest in tst to match prod environment 
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
        missing_test_types = self.missing_test_report(tst_df,prod_df, prodTestsMissingInTST)

        # going to filter out non matches of accession number on new findings 
        missing_report = prod_df[~prod_df['DILR_AccessionNumber'].isin(new_findings['DILR_AccessionNumber'])]
        
        # Seeing if there are any ResultTests that I said are missing, which are actually in TST  
        final_missing = self.sanity_check(prod_df, tst_df, missing_report)

        return one2one_df, final_missing, missing_test_types

    def report_builder(self, prod_df, tst_df):
        """
        Generates the validation reports for the given production and test dataframes.
        
        This function takes two pandas DataFrames, `prod_df` and `tst_df`, representing
        the production and test data respectively. It performs the following steps:
        
        1. Calls the `final_summary_reports` method to obtain a list of initial missing
        reports, tests seen in PROD but not TST, and a summary of tests that were
        matched in TST but not in PROD.
        
        2. Builds an Excel sheet named after the `lab_name` attribute of the class,
        replacing any non-alphanumeric characters with underscores, and creates a
        `pandas.ExcelWriter` object.
        
        3. Writes the list of missing test types to the Excel sheet, in a sheet named
        "Missing_TestTypes", starting from the first column.
        
        4. Writes the one-to-one matching summary to the Excel sheet, in a sheet named
        "One2One_matches", starting from the first column.
        
        5. Writes the final list of missing reports to the Excel sheet, in a sheet named
        "Tests in Prod not TST", starting from the first column.
        
        6. Closes the Excel writer object.
        
        Args:
            prod_df (pandas.DataFrame): The production dataframe containing the reports from PROD environment.
            tst_df (pandas.DataFrame): The test dataframe containing the reports frmo TST environment.
        
        Returns:
            None
        Raises:
            None
        """

        # getting list of initial missing reports, tests seen in PROD but not TST
        # and a summary of tests that were matched in TST but not prod 
        one2one_df, final_missing, missing_test_types = self.final_summary_reports(prod_df, tst_df)

        # building excel sheet with all reports
        report_name = re.sub(r'[^\w\s]+', '_',self.lab_name)
        writer = pd.ExcelWriter(f'{report_name}_validation_reports.xlsx', engine='xlsxwriter')

        # list of missing test types
        missing_test_types.to_excel(
            writer, 
            sheet_name='Missing_TestTypes', 
            startcol=0, 
            index=False
            )
        
        # Matching summary
        one2one_df.to_excel(
            writer, 
            sheet_name='One2One_matches', 
            startcol=0, 
            index=False
            )
        
        # final missing reports list
        final_missing.to_excel(
            writer, 
            sheet_name='Tests in Prod not TST', 
            startcol=0, 
            index=False
            )
        writer.close()

        return 