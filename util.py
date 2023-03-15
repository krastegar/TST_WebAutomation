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
import pyodbc
import os
import rapidfuzz
from rapidfuzz import fuzz

def tstRangeQuery_lab(test_center_1, test_center_2 = None):
    '''
    Query string that is meant to extract relative information for looking at TST WebCMR data
    in bulk. In this example I am only using two test centers. If we are looking at 3 or more then
    I can make the requirements for the function a list or something.
    '''
    query = f'''
    SELECT 
        ACCESSIONNUMBER,
        ORDERRESULTSTATUS, 
        OBSERVATIONRESULTSTATUS,
        SPECCOLLECTEDDATE,
        SPECRECEIVEDDATE, 
        RESULTDATE,
        TESTCODE, 
        RESULTTEXT, 
        OrganismCode,
        ResultedOrganism,
        ABNORMALFLAG,
        REFERENCERANGE,
        SPECIMENSOURCE,
        PROVIDERNAME,
        PROVIDERADDRESS,
        PROVIDERCITY,
        PROVIDERSTATE,
        PROVIDERZIP,
        PROVIDERPHONE,
        FACILITYADDRESS,
        FACILITYCITY,
        FACILITYSTATE,
        FACILITYZIP,
        FACILITYPHONE, 
        FACILITYNAME
    FROM 
        [Laboratory Information (system)]
    WHERE 
        FACILITYNAME LIKE '%{test_center_1}%' OR FACILITYNAME LIKE '%{test_center_2}%'
    '''
    return query 

def tstRangeQuery_demographic(test_center_1, test_center_2 = None):
    '''
    Similar objective to tstRangeQuery_lab, but the select statement is meant
    to get different information from the disease incident table from the .accdb file
    '''
    
    query = f'''
    SELECT 
        Last_Name,
        First_Name, 
        DOB, 
        Street_Address,
        City,
        State,
        Zip,
        Home_Telephone, 
        Reported_Race as Race,
        Ethnicity
    FROM 
        [Disease Incident Export]
    WHERE 
        Laboratory LIKE '%{test_center_1}%' OR Laboratory LIKE '%{test_center_2}%'
        '''
    
    return query

def completeness_report(query, conn):
    '''
    After constructing the query and making the connection to the database. We create 
    a dataframe that summarize the results of the bulk exports 
    '''
    df = pd.read_sql_query(query, conn)
    # df.to_excel(r'C:\Users\krastega\Desktop\MicrosoftAccessDB\DI.xlsx')
    
    # Getting counts of Null and not Null values
    null_counts = df.isna().sum()
    nonNullCounts = df.count()
    
    # Calculating Percentage of complete information from those values. 
    difCounts = np.absolute(null_counts.values-nonNullCounts.values)
    total_num = null_counts.values+nonNullCounts.values
    percent_complete = (difCounts/total_num)*100

    # Lab df done 
    lab_df = pd.DataFrame(
        {
        'Fields of Interest': list(df.columns),
        'Percent Complete' : percent_complete
        }
    )
    lab_df['Percent Complete'] = lab_df['Percent Complete'].map('{:,.1f}'.format)

    return lab_df

def database_connection(file_name, folder_path):
    '''
    Creating database connection to run sql queries and extract necessary
    information from Microsoft Access files
    '''
    pyodbc.lowercase = False
    conn = pyodbc.connect(
        r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" +
        fr"Dbq={folder_path}\{file_name}.accdb")
    cursor = conn.cursor()

    return (conn, cursor)

def combined_MM(folder_path, startswith):
    '''
    Function that looks into a directory and looks for excel files that start with 'startwith' 
    variable. Grabs all of the file names and reads them into pandas df, which is later 
    concatenated 
    '''
    # getting list of files for either TST or Prod IMM exports
    prod_file_list = [os.path.join(folder_path, file) for file in os.listdir(folder_path)
                      if os.path.basename(file).startswith(startswith) and os.path.basename(file).endswith('.xlsx')]
    
    # combining all of the excel files into 1 df
    prod_db_array = []
    for file in prod_file_list: 
        df = pd.read_excel(file)
        prod_db_array.append(df)
    
    combined_df = pd.concat(prod_db_array)
    return combined_df

def test_match(df, column1, column2):
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
    '''
    new_df structure:

                                                SarsCov2        HepC SurfAnti A
    SarsCoronaVirusLike::PROBE:ANT                  15               0

    HepatitusC Surface Antingen::MOLC:ANTH          0                42

    Code below is looping through each row and finding the column with the highest counts, if there
    are 2 columns that have the same amount of counts it looks at both columns to see if there are 
    any other rows within that column that have higher counts. If there are then it removes the column
    name from the list of tie breaker columns 
    '''
    new_df.to_excel('Intermediate_match.xlsx')
    matched_pairs = {}
    match_array = []
    for index, row in new_df.iterrows():
        max_val = row.max()
        col_names = row[row == max_val].index.tolist()
        if len(col_names) > 1: 
            for col in col_names: 
                tie_break_score = new_df[col].max()
                if tie_break_score > max_val: 
                    col_names.remove(new_df[col].name)
        matched_pairs[col_names[0]] = index
        match_array.append((index, col_names[0]))
    match_df = pd.DataFrame(match_array, columns=['ProdTest', 'TSTTest'])
    
    # Calculating string similarity score using fuzzywuzzy and 
    # this will later be used to filter out duplicate matches
    match_df['Ratio_score'] = match_df.apply(
        lambda x: fuzz.token_ratio(x['ProdTest'], x['TSTTest']), 
        axis=1
        )
    match_df.to_excel('match.xlsx')

    return matched_pairs

