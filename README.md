# Excel Report Generator

## Overview
This app generates multiple reports based on an input Excel file.

## How to Use
1. Place your `input_file.xlsx` in the same directory as the app.
2. Run the executable:
   - On Windows:
     ```bash
     main.exe
     ```
   - On Mac/Linux:
     ```bash
     ./main
     ```
3. The `output_file.xlsx` will be generated in the same directory.

## Requirements
- Python 3.7 or higher (for running the source code)
- Required libraries:
  - `pandas`
  - `openpyxl`
  - `xlsxwriter`

## Modules
The app includes the following modules:
- Invoice Summary
- Route-wise Summary
- SKU Summary
- Sales Analytics
- Daily Sales Summary
- Average SKU per Invoice

## Testing
Use a sample `input_file.xlsx` with the required format for testing.
