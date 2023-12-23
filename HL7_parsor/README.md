# TST WebCMR Validation Automation
## Purpose
The TST WebCMR Validation Automation script is designed to automate the validation of TST WebCMR data using HL7 messages received from a specific laboratory. The script performs the following tasks:

## Procedure
### 1. Login to TST:

    Provides credentials to log in to the TST system.

### 2. Search for Lab:

    Searches for the specified laboratory within the TST system.

### 3. Export Incoming Messages:

    Exports incoming messages from the specified laboratory in bulk.

### 4. Classify Cases:

    Classifies cases into three categories: Categorical, Numeric, Mixed.

### 5. Scrape TSTWebCMR Sections:

    Scrapes TSTWebCMR Demographics and Laboratory sections for specific cases.

### 6. Validate with HL7 Messages:

    Compares sections in the HL7 message with the corresponding sections in TSTWebCMR.

### 7. Produce Summary Excel Report:

    Generates a summary Excel report based on the validation results.

### 8. Logout:

    Logs out of the TST system.

## Usage
### 1. Installation:

        Ensure you have the correct version of Chrome installed.
        Place the compatible chromedriver in the same folder as the script.

### 2. Configuration:

        Open the script and input the required configuration parameters:
        TST login credentials.
        Laboratory of interest.
        Date range for validation.

### 3. Execution:

        Run the script.

## Troubleshooting
If you encounter issues with the script, consider the following steps:

### 1. ChromeDriver Compatibility:

        Ensure the chromedriver in the script folder matches your Chrome version.
        Replace the chromedriver with a compatible version if needed.

### 2. TST Login Credentials:

        Verify that the provided TST login credentials are accurate.
        Update credentials in the script if necessary.

### 3. Laboratory and Date Range:

        Double-check the specified laboratory and date range in the script.
        Ensure the specified laboratory exists in the TST system.

### 4. Error Messages:

        Check the console for any error messages or logs generated during execution.

### 5. File Output:

        Ensure the script has permission to write files in the specified directory.

## Note:
Remember to follow responsible usage practices and adhere to any terms of use associated with the TST system. Use the script in accordance with applicable laws and regulations.