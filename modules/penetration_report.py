# import pandas as pd

# # Constants for column names and sheet names
# MOVEMENTS_SHEET = "movements"
# PRODUCTS_SHEET = "products"
# OUTLETS_SHEET = "outlets"
# REQUIRED_COLUMNS = {
#     MOVEMENTS_SHEET: ["Code", "Quantity", "OutletCode"],
#     PRODUCTS_SHEET: ["Code", "SKUDescription"],
#     OUTLETS_SHEET: ["OutletCode", "RouteCode"],
# }

# def validate_columns(df, sheet_name):
#     """
#     Validates if the required columns are present in the dataframe.
    
#     Args:
#         df (DataFrame): DataFrame to check for required columns.
#         sheet_name (str): The sheet name for logging error messages.
    
#     Returns:
#         bool: True if all required columns are present, otherwise False.
#     """
#     missing_columns = [col for col in REQUIRED_COLUMNS[sheet_name] if col not in df.columns]
#     if missing_columns:
#         print(f"Error: Missing columns in '{sheet_name}' sheet: {missing_columns}")
#         return False
#     return True

# def generate_penetration_report(input_file, writer):
#     """
#     Generates the Penetration Report and writes it to the output Excel file.
    
#     Args:
#         input_file (str): Path to the input Excel file.
#         writer (pd.ExcelWriter): Excel writer object to write the output file.
#     """
#     print("Generating Penetration Report...")
#     try:
#         print("Loading sheets from input file...")
#         movements = pd.read_excel(input_file, sheet_name=MOVEMENTS_SHEET)
#         products = pd.read_excel(input_file, sheet_name=PRODUCTS_SHEET)
#         outlets = pd.read_excel(input_file, sheet_name=OUTLETS_SHEET)
#     except ValueError as e:
#         print(f"Error reading input file: {e}")
#         return
#     except Exception as e:
#         print(f"An unexpected error occurred while loading sheets: {e}")
#         return
    
#     # Validate required columns in each sheet
#     for sheet_name, df in zip([MOVEMENTS_SHEET, PRODUCTS_SHEET, OUTLETS_SHEET], 
#                               [movements, products, outlets]):
#         if not validate_columns(df, sheet_name):
#             return

#     # Merge movements with products and outlets
#     try:
#         print("Merging movements with products and outlets...")
#         merged_data = pd.merge(movements, products[["Code", "SKUDescription"]], how="left", on="Code")
#         merged_data = pd.merge(merged_data, outlets[["OutletCode", "RouteCode"]], how="left", on="OutletCode")
#     except KeyError as e:
#         print(f"Error during merging data: Missing key {e}")
#         return
#     except Exception as e:
#         print(f"An unexpected error occurred during data merging: {e}")
#         return

#     # Check for missing RouteCode values and fill them
#     print("Missing RouteCode values:", merged_data["RouteCode_y"].isnull().sum())
#     merged_data["RouteCode_y"].fillna('Unknown', inplace=True)

#     # Filter data to get outlets with non-zero sales for each SKU
#     merged_data["Sold"] = merged_data["Quantity"] > 0

#     # Remove duplicate OutletCode for each SKU to count unique outlets
#     unique_outlets = merged_data.drop_duplicates(subset=["SKUDescription", "RouteCode_y", "OutletCode"])

#     # Create a pivot table to count unique outlets with sales per SKU for each route
#     try:
#         print("Creating pivot table for Penetration Report...")
#         pivot_table = unique_outlets.pivot_table(
#             index="RouteCode_y",
#             columns="SKUDescription",  # Assuming SKUDescription is the desired column
#             values="Sold",
#             aggfunc="sum",
#             fill_value=0,
#         )
#     except Exception as e:
#         print(f"An error occurred while creating the pivot table: {e}")
#         return

#     # Reset index for tabular format
#     penetration_report = pivot_table.reset_index()

#     # Calculate percentages based on the total for each SKU
#     percentage_report = penetration_report.copy()
#     for column in penetration_report.columns[1:]:  # Skip RouteCode_y column
#         percentage_report[column] = (penetration_report[column] / penetration_report[column].sum()) * 100

#     # Round the percentage values to one decimal place
#     percentage_report = percentage_report.round(1)

#     # Debugging: Check the generated reports
#     print("Generated Numeric Penetration Report Data:")
#     print(penetration_report.head())
#     print("Generated Percentage Penetration Report Data:")
#     print(percentage_report.head())

#     # Write the reports to the Excel file
#     try:
#         print("Writing the Numeric and Percentage Penetration Reports to the Excel file...")
#         penetration_report.to_excel(writer, sheet_name="Penetration Numeric", index=False, startrow=1)
#         percentage_report.to_excel(writer, sheet_name="Penetration Percentage", index=False, startrow=1)
#     except Exception as e:
#         print(f"An error occurred while writing the reports to the Excel file: {e}")
#         return

#     # Format the sheets
#     try:
#         print("Formatting the Excel sheets...")
#         workbook = writer.book
        
#         # Numeric Report Sheet
#         worksheet_numeric = writer.sheets["Penetration Numeric"]
#         header_format = workbook.add_format({
#             "bold": True,
#             "text_wrap": True,
#             "valign": "center",
#             "align": "center",
#             "fg_color": "#D7E4BC",
#             "border": 1,
#         })
#         for col_num, value in enumerate(penetration_report.columns.values):
#             worksheet_numeric.write(0, col_num, value, header_format)
        
#         # Percentage Report Sheet
#         worksheet_percentage = writer.sheets["Penetration Percentage"]
#         for col_num, value in enumerate(percentage_report.columns.values):
#             worksheet_percentage.write(0, col_num, value, header_format)

#         # Adjust column widths dynamically for both sheets
#         for worksheet in [worksheet_numeric, worksheet_percentage]:
#             for col_num, col_name in enumerate(penetration_report.columns):
#                 column_width = max(
#                     penetration_report[col_name].astype(str).apply(len).max(),
#                     len(col_name) + 2  # Add some padding
#                 )
#                 worksheet.set_column(col_num, col_num, column_width)
#     except Exception as e:
#         print(f"An error occurred while formatting the Excel sheets: {e}")
#         return

#     print("Penetration Report generation complete.")

# if __name__ == "__main__":
#     input_file = "input_file.xlsx"
#     output_file = "output_file.xlsx"
#     try:
#         with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
#             print("Generating the Penetration Report...")
#             generate_penetration_report(input_file, writer)
#     except Exception as e:
#         print(f"Error during penetration report generation: {e}")

import pandas as pd

# Constants for column names and sheet names
MOVEMENTS_SHEET = "movements"
PRODUCTS_SHEET = "products"
OUTLETS_SHEET = "outlets"
REQUIRED_COLUMNS = {
    MOVEMENTS_SHEET: ["Code", "Quantity", "OutletCode"],
    PRODUCTS_SHEET: ["Code", "SKUDescription"],
    OUTLETS_SHEET: ["OutletCode", "RouteCode"],
}

def validate_columns(df, sheet_name):
    """
    Validates if the required columns are present in the dataframe.
    
    Args:
        df (DataFrame): DataFrame to check for required columns.
        sheet_name (str): The sheet name for logging error messages.
    
    Returns:
        bool: True if all required columns are present, otherwise False.
    """
    missing_columns = [col for col in REQUIRED_COLUMNS[sheet_name] if col not in df.columns]
    if missing_columns:
        print(f"Error: Missing columns in '{sheet_name}' sheet: {missing_columns}")
        return False
    return True

def generate_penetration_report(input_file, writer):
    """
    Generates the Penetration Report and writes it to the output Excel file.
    
    Args:
        input_file (str): Path to the input Excel file.
        writer (pd.ExcelWriter): Excel writer object to write the output file.
    """
    print("Generating Penetration Report...")
    try:
        print("Loading sheets from input file...")
        movements = pd.read_excel(input_file, sheet_name=MOVEMENTS_SHEET)
        products = pd.read_excel(input_file, sheet_name=PRODUCTS_SHEET)
        outlets = pd.read_excel(input_file, sheet_name=OUTLETS_SHEET)
    except ValueError as e:
        print(f"Error reading input file: {e}")
        return
    except Exception as e:
        print(f"An unexpected error occurred while loading sheets: {e}")
        return
    
    # Validate required columns in each sheet
    for sheet_name, df in zip([MOVEMENTS_SHEET, PRODUCTS_SHEET, OUTLETS_SHEET], 
                              [movements, products, outlets]):
        if not validate_columns(df, sheet_name):
            return

    # Merge movements with products and outlets
    try:
        print("Merging movements with products and outlets...")
        merged_data = pd.merge(movements, products[["Code", "SKUDescription"]], how="left", on="Code")
        merged_data = pd.merge(merged_data, outlets[["OutletCode", "RouteCode"]], how="left", on="OutletCode")
    except KeyError as e:
        print(f"Error during merging data: Missing key {e}")
        return
    except Exception as e:
        print(f"An unexpected error occurred during data merging: {e}")
        return

    # Check for missing RouteCode values and fill them
    print("Missing RouteCode values:", merged_data["RouteCode_y"].isnull().sum())
    merged_data["RouteCode_y"].fillna('Unknown', inplace=True)

    # Filter data to get outlets with non-zero sales for each SKU
    merged_data["Sold"] = merged_data["Quantity"] > 0

    # Remove duplicate OutletCode for each SKU to count unique outlets
    unique_outlets = merged_data.drop_duplicates(subset=["SKUDescription", "RouteCode_y", "OutletCode"])

    # Create a pivot table to count unique outlets with sales per SKU for each route
    try:
        print("Creating pivot table for Penetration Report...")
        pivot_table = unique_outlets.pivot_table(
            index="RouteCode_y",
            columns="SKUDescription",  # Assuming SKUDescription is the desired column
            values="Sold",
            aggfunc="sum",
            fill_value=0,
        )
    except Exception as e:
        print(f"An error occurred while creating the pivot table: {e}")
        return

    # Reset index for tabular format
    penetration_report = pivot_table.reset_index()

    # Calculate total outlets per RouteCode from the "Outlets" sheet
    total_outlets_per_route = outlets.groupby("RouteCode")["OutletCode"].nunique()

    # Create a copy for the percentage report
    percentage_report = penetration_report.copy()

    # Calculate percentages based on the total outlets per route
    for column in penetration_report.columns[1:]:  # Skip RouteCode_y column
        percentage_report[column] = (
            penetration_report[column] / penetration_report["RouteCode_y"].map(total_outlets_per_route)
        ) * 100

    # Round the percentage values to one decimal place
    percentage_report = percentage_report.round(1)

    # Debugging: Check the generated reports
    print("Generated Numeric Penetration Report Data:")
    print(penetration_report.head())
    print("Generated Percentage Penetration Report Data:")
    print(percentage_report.head())

    # Write the reports to the Excel file
    try:
        print("Writing the Numeric and Percentage Penetration Reports to the Excel file...")
        penetration_report.to_excel(writer, sheet_name="Penetration Numeric", index=False, startrow=1)
        percentage_report.to_excel(writer, sheet_name="Penetration Percentage", index=False, startrow=1)
    except Exception as e:
        print(f"An error occurred while writing the reports to the Excel file: {e}")
        return

    # Format the sheets
    try:
        print("Formatting the Excel sheets...")
        workbook = writer.book
        
        # Numeric Report Sheet
        worksheet_numeric = writer.sheets["Penetration Numeric"]
        header_format = workbook.add_format({
            "bold": True,
            "text_wrap": True,
            "valign": "center",
            "align": "center",
            "fg_color": "#D7E4BC",
            "border": 1,
        })
        for col_num, value in enumerate(penetration_report.columns.values):
            worksheet_numeric.write(0, col_num, value, header_format)
        
        # Percentage Report Sheet
        worksheet_percentage = writer.sheets["Penetration Percentage"]
        for col_num, value in enumerate(percentage_report.columns.values):
            worksheet_percentage.write(0, col_num, value, header_format)

        # Adjust column widths dynamically for both sheets
        for worksheet in [worksheet_numeric, worksheet_percentage]:
            for col_num, col_name in enumerate(penetration_report.columns):
                column_width = max(
                    penetration_report[col_name].astype(str).apply(len).max(),
                    len(col_name) + 2  # Add some padding
                )
                worksheet.set_column(col_num, col_num, column_width)
    except Exception as e:
        print(f"An error occurred while formatting the Excel sheets: {e}")
        return

    print("Penetration Report generation complete.")

if __name__ == "__main__":
    input_file = "input_file.xlsx"
    output_file = "output_file.xlsx"
    try:
        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            print("Generating the Penetration Report...")
            generate_penetration_report(input_file, writer)
    except Exception as e:
        print(f"Error during penetration report generation: {e}")
