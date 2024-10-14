# SO Data Processing Macro

This project involves a comprehensive macro designed to automate the extraction, transformation, and validation of *Sales Order (SO)* data within an Excel workbook. The macro connects and consolidates data from multiple workbooks and sheets, filters the data, and performs tolerance checks, significantly streamlining the monthly data processing tasks.

## Macro Overview

The macro streamlines the process of handling SO data across various workbooks, including:
- **KW PAEN SO Final Workbook**
- **KW RA Working Workbook**

The macro performs the following key tasks:

1. **Data Preparation**:
    - Prompts the user to input the **month** and **RA stage**.
    - Clears and prepares the relevant ranges in the **FinalBook** (Sales Order Final Workbook).
    - Copies and pastes data from the **SourceBook** (Sales Order Raw Workbook) into the appropriate worksheets of the **FinalBook**.
    - Applies auto-filling and formats to ensure that the newly inserted data is consistent.

2. **Data Filtering and Copying**:
    - Applies filters to the *Sales Order Data* sheet, retrieving only the **RELEASED** and **TECHNICALLY COMPLETED** entries.
    - Clears existing data in the **SO Data** sheet of the FinalBook and copies the filtered data from the **SourceBook**.
    - Replaces specific negative values with 0 for certain sheets and highlights them with a green background.

3. **Pivot Table Refresh**:
    - Refreshes the pivot tables for both **Periodic** and **Cumulative** data.
    - Copies the refreshed data from the pivot tables and pastes it into the corresponding sheets, updating the values in the **FinalBook**.

4. **Variance Checks**:
    - Filters data in the **Cumulative Data** sheet for **RELEASED** entries and excludes zero values.
    - Clears and prepares the **Variance** sheet for new data.
    - Copies specific columns of data from the **Cumulative Data** sheet to the **Variance** sheet for further comparison.
    - Applies tolerance checks to identify discrepancies in the data.
    - Automatically deletes rows with differences smaller than the tolerance level.

5. **Sorting and Final Formatting**:
    - Applies a final sort on the **Variance** sheet based on the identified discrepancies.
    - Completes the process by refreshing all data connections and presenting a message box indicating completion.

## Key Features

- **Data Consolidation**: The macro automates the process of copying, pasting, and formatting data from various sources into the main file, ensuring the information is up-to-date.
- **Pivot Table Refresh**: Automatically updates and refreshes both periodic and cumulative pivot tables to provide accurate summary data.
- **Variance and Tolerance Checks**: Ensures data accuracy by checking for variances and applying tolerance levels to filter out insignificant discrepancies.
- **Custom User Input**: Prompts the user for specific month and period details to tailor the data processing to the correct timeframe.
