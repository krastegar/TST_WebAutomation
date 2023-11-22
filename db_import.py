import pyodbc
import pandas as pd 
import numpy as np

# connecting python to microsoft access file
pyodbc.lowercase = False
conn = pyodbc.connect(
    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" +
    r"Dbq=C:\Users\krastega\Desktop\MicrosoftAccessDB\TST_DIE_02132023_02212023_exp022323.accdb")
cursor = conn.cursor()

# looking at Disease Incidents table specifically and reading it into a dataframe

query = '''
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
    FACILITYNAME LIKE '%Palomar%' OR FACILITYNAME LIKE '%None%'
'''
# LIKE % String % is a wild card operator that looks for the substring 'String' in the facility names
df = pd.read_sql_query(query, conn)
df.to_excel(r'C:\Users\krastega\Desktop\MicrosoftAccessDB\DI.xlsx')
null_counts = df.isna().sum()
nonNullCounts = df.count()
difCounts = np.absolute(null_counts.values-nonNullCounts.values)
total_num = null_counts.values+nonNullCounts.values
percent_complete = (difCounts/total_num)*100

# Lab df done 
lab_df = pd.DataFrame(percent_complete, index=null_counts.index, columns=['Percent Complete'])
lab_df = lab_df['Percent Complete'].map('{:,.1f} %'.format)
print(lab_df)


