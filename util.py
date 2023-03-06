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
    lab_df = pd.DataFrame(percent_complete, index=null_counts.index, columns=['Percent Complete'])
    lab_df = lab_df['Percent Complete'].map('{:,.1f}'.format)

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