# Bulk Incoming Message Monitor Validation App
## Overview
    The WebCMR HL7 Mapping App is designed to analyze and map HL7 information from incoming message monitor (IMM) exports in two different WebCMR environments: PROD and TST. The app performs the following steps to create a comprehensive report:

### 1. Combine IMM Exports:

    Takes IMM exports from a specified folder, combines them, and separates them into PROD and TST worksheets.

### 2. Create Resulted Test Mapping Dictionary:

    Matches resulted test columns from both PROD and TST worksheets based on the accession number.
    Identifies frequently matched tests by accession number to create an equivalence mapping structure.
    Transforms TST resulted test names to match the names in the PROD environment using the mapping structure.

### 3. Filter and Classify Missing Tests:

    Performs a second join on resulted test name and accession number.
    Filters out tests not seen in the newly matched dataframe, classifying them as missing.

### 4. Generate Comprehensive Report:

    Produces an Excel worksheet with three tabs:
        One-to-one matching key between PROD and TST.
        Test types seen in PROD but not in TST, and vice versa.
        A comprehensive list of HL7 file names seen in PROD but not in TST.

## Usage
### 1. Run the App:

    Double-click on the executable file (webcmr_mapping_app.exe) to launch the application.

### 2. Input Dialog:

    The application will prompt you with a dialog box asking for the folder path containing IMM exports and the desired report name.

### 3. Provide Input:

    Enter the folder path where IMM exports are stored.
    Enter the desired name for the final report.

### 4. Generate Report:

    After input information is received the program will generate the report automatically and 
    close after completed. 
        Note: If the there is not any report card produced and the error_log does not show "program complete" message than there was an issue. Please look at the traceback to debug source code
    
## Disclaimer
IMM exports must start with "TST" for the TST environment and "PROD" for the PROD environment.