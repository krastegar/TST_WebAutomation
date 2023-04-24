#-------------------------------------------------------------------------------------------
# 
#   Purpose:
#   Creating a 'Completeness Report' based on Range Exports from TST environment. This class 
#   will look at specific fields in the TST.accdb file and calculate the percentage of how many 
#   non-null / total entries seen in these specific fields. These fields are similar to what we see 
#   on the demographic and lab tabs of WebCMR. 
#   
#   Algorithm:
#       1. Create a database connection to .accdb file
#       2. Construct 2 queries. One to get information related to Demographic tab 
#          and another for information from Laboratory tab
#       3. Calculate the total of non_missing values over the complete total for that specific field
#          for all fields
#       4. Put values into a dataframe with percentages and export it as a .xlsx file
#       5. Done 
#
#   Author: Kiarash Rastegar
#   Date: 4/19/23
#-------------------------------------------------------------------------------------------

import pandas as pd
import numpy as np
import pyodbc
import re 

class Completeness:
    def __init__(
            self,
            lab_name,
            file_name,
            folder_path,
            test_center_1, 
            test_center_2 = None,
            test_center_3 = None,
            test_center_4 = None,
            test_center_5 = None
    ):
        self.lab_name = lab_name
        self.file_name = file_name
        self.folder_path = folder_path
        self.test1 = test_center_1
        self.test2 = test_center_2
        self.test3 = test_center_3
        self.test4 = test_center_4
        self.test5 = test_center_5

    def database_connection(self):
        '''
        Creating database connection to run sql queries and extract necessary
        information from Microsoft Access files
        '''
        pyodbc.lowercase = False
        conn = pyodbc.connect(
            r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" +
            fr"Dbq={self.folder_path}\{self.file_name}.accdb")
        cursor = conn.cursor()

        return conn, cursor
    
    def tstRangeQuery_lab(self):
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
            FACILITYNAME LIKE '%{self.test1}%' 
            OR FACILITYNAME LIKE '%{self.test2}%'
            OR FACILITYNAME LIKE '%{self.test3}%'
            OR FACILITYNAME LIKE '%{self.test4}%'
            OR FACILITYNAME LIKE '%{self.test5}%'
        '''
        return query
    
    def tstRangeQuery_demographic(self):
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
            Laboratory LIKE '%{self.test1}%' 
            OR Laboratory LIKE '%{self.test2}%'
            OR Laboratory LIKE '%{self.test3}%'
            OR Laboratory LIKE '%{self.test4}%'
            OR Laboratory LIKE '%{self.test5}%'
            '''
        return query

    def completeness_report(self):
        '''
        After generating queries used to grab information from the .accdb files that are related to 
        WebCMR Lab and Demographics tab, we create a dataframe from them using range_export() method.
        The dataframe contains percent completeness of each desired field. Finally we combine both 
        dataframes into one excel sheet, which results in our final completeness summary report
        '''
        
        # need to establish database connection
        conn, _ = self.database_connection()

        # generating queries for both Laboratory data and Demographic data 
        lab_query = self.tstRangeQuery_lab()
        demo_query = self.tstRangeQuery_demographic()

        # Creating dataframes from query results 
        lab_df = self.range_export_df(conn, lab_query)
        demo_df = self.range_export_df(conn, demo_query)

        # Going to make one excel sheet with completeness reports from both demographic 
        # and lab information
        lab_name = re.sub(r'[^\w\s]+', '_',self.lab_name)
        writer = pd.ExcelWriter(f'{lab_name}_completeness_reports.xlsx', engine='xlsxwriter')
        demo_df.to_excel(writer, sheet_name='Sheet1', startcol=0, index=False)
        lab_df.to_excel(writer, sheet_name='Sheet1', startcol = len(demo_df.columns)+1, index=False)
        
        # close writer object
        writer.close()
        return 

    def range_export_df(self, conn, query):
        '''
        After constructing the query and making the connection to the database. We create 
        a dataframe that summarize the results of the bulk exports. The summary is done by looking
        at the total amount of NonNull values / total (NonNull + Null) for each specified field in 
        the query that is passed to this method 
        '''
        df = pd.read_sql_query(query, conn)
        
        # Getting counts of Null and not Null values
        null_counts = df.isna().sum()
        nonNullCounts = df.count()
        
        # Calculating Percentage of complete information from those values. 
        difCounts = np.absolute(nonNullCounts.values) # this is a difference of arrays (might just want a ratio of )
        total_num = null_counts.values+nonNullCounts.values
        percent_complete = (difCounts/total_num)*100

        # Lab df done 
        lab_df = pd.DataFrame(
            {
            'Fields of Interest': list(df.columns),
            'Percent Complete' : percent_complete
            }
        )
        lab_df['Percent Complete'] = lab_df['Percent Complete'].map('{:,.2f}'.format)
        return lab_df