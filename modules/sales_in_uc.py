# import pandas as pd

# def generate_sales_in_uc_report(input_file, writer):
#     """
#     Generates the Sales in UC report (Pack Wise and Brand Wise) and writes them to the output Excel file.

#     Args:
#         input_file (str): Path to the input Excel file.
#         writer (pd.ExcelWriter): Excel writer object to write the output file.
#     """
#     try:
#         print("Loading movements and products sheets...")
#         movements = pd.read_excel(input_file, sheet_name="movements")
#         products = pd.read_excel(input_file, sheet_name="products")
#     except ValueError as e:
#         print("Error loading the sheets. Ensure the input file contains 'movements' and 'products'.")
#         raise ValueError("The input file must contain 'movements' and 'products' sheets.") from e
#     except Exception as e:
#         print(f"An error occurred while loading the sheets: {e}")
#         raise

#     # Merge movements with products on "Code" (SKU Code)
#     try:
#         print("Merging movements with products on 'Code' (SKU Code)...")
#         merged_data = pd.merge(
#             movements,
#             products[["Code", "Pack", "Brand", "Unitfactor"]],
#             how="left",
#             left_on="Code",
#             right_on="Code",
#         )
#     except Exception as e:
#         print(f"An error occurred while merging data: {e}")
#         raise

#     # Warn if there are unmatched product codes
#     unmatched_codes = merged_data[merged_data["Pack"].isna()]
#     if not unmatched_codes.empty:
#         print(f"Warning: {len(unmatched_codes)} product codes in 'movements' were not found in 'products'.")

#     # Adjust quantity by multiplying with the Unitfactor
#     merged_data["AdjustedQuantity"] = merged_data["Quantity"] * merged_data["Unitfactor"]

#     # Generate the Pack Wise Sale in UC report
#     try:
#         print("Creating the pivot table for Pack Wise Sale in UC Report...")
#         pack_pivot = merged_data.pivot_table(
#             index="RouteCode",
#             columns="Pack",
#             values="AdjustedQuantity",
#             aggfunc="sum",
#             fill_value=0,
#         )
#         # Add a "Total" row at the end of the pack pivot table
#         pack_total_row = pack_pivot.sum(numeric_only=True, skipna=True)
#         pack_total_row.name = "Total"
#         pack_pivot = pd.concat([pack_pivot, pack_total_row.to_frame().T])
#     except Exception as e:
#         print(f"An error occurred while creating the pack pivot table: {e}")
#         raise

#     # Reset the index to make it tabular
#     pack_wise_report = pack_pivot.reset_index()

#     # Write the Pack Wise Sale in UC report to the Excel file
#     try:
#         print("Writing the Pack Wise Sale in UC Report to the Excel sheet...")
#         pack_wise_report.to_excel(writer, sheet_name="Pack Wise Sale in UC", index=False, startrow=1)
#     except Exception as e:
#         print(f"An error occurred while writing the report to the Excel file: {e}")
#         raise

#     # Format the Pack Wise Sale in UC sheet
#     try:
#         print("Formatting the Pack Wise Sale in UC Excel sheet...")
#         workbook = writer.book
#         worksheet = writer.sheets["Pack Wise Sale in UC"]

#         worksheet.write(0, 0, "Pack Wise Sale in UC Report")

#         header_format = workbook.add_format({
#             "bold": True,
#             "text_wrap": True,
#             "valign": "center",
#             "align": "center",
#             "fg_color": "#D7E4BC",
#             "border": 1,
#         })

#         for col_num, value in enumerate(pack_wise_report.columns.values):
#             worksheet.write(1, col_num, value, header_format)

#         for col_num, col_name in enumerate(pack_wise_report.columns):
#             max_width = max(pack_wise_report[col_name].astype(str).apply(len).max(), len(str(col_name))) + 2
#             worksheet.set_column(col_num, col_num, max_width)
#     except Exception as e:
#         print(f"An error occurred while formatting the Excel sheet: {e}")
#         raise

#     # Generate the Brand Wise Sale in UC report
#     try:
#         print("Creating the pivot table for Brand Wise Sale in UC Report...")
#         brand_pivot = merged_data.pivot_table(
#             index="RouteCode",
#             columns="Brand",
#             values="AdjustedQuantity",
#             aggfunc="sum",
#             fill_value=0,
#         )
#         # Add a "Total" row at the end of the brand pivot table
#         brand_total_row = brand_pivot.sum(numeric_only=True, skipna=True)
#         brand_total_row.name = "Total"
#         brand_pivot = pd.concat([brand_pivot, brand_total_row.to_frame().T])
#     except Exception as e:
#         print(f"An error occurred while creating the brand pivot table: {e}")
#         raise

#     # Reset the index to make it tabular
#     brand_wise_report = brand_pivot.reset_index()

#     # Write the Brand Wise Sale in UC report to the Excel file
#     try:
#         print("Writing the Brand Wise Sale in UC Report to the Excel sheet...")
#         brand_wise_report.to_excel(writer, sheet_name="Brand Wise Sale in UC", index=False, startrow=1)
#     except Exception as e:
#         print(f"An error occurred while writing the report to the Excel file: {e}")
#         raise

#     # Format the Brand Wise Sale in UC sheet
#     try:
#         print("Formatting the Brand Wise Sale in UC Excel sheet...")
#         worksheet = writer.sheets["Brand Wise Sale in UC"]

#         worksheet.write(0, 0, "Brand Wise Sale in UC Report")

#         for col_num, value in enumerate(brand_wise_report.columns.values):
#             worksheet.write(1, col_num, value, header_format)

#         for col_num, col_name in enumerate(brand_wise_report.columns):
#             max_width = max(brand_wise_report[col_name].astype(str).apply(len).max(), len(str(col_name))) + 2
#             worksheet.set_column(col_num, col_num, max_width)
#     except Exception as e:
#         print(f"An error occurred while formatting the Excel sheet: {e}")
#         raise

#     print("Sales in UC Report generation complete.")

# if __name__ == "__main__":
#     try:
#         with pd.ExcelWriter("output_file.xlsx", engine="xlsxwriter") as writer:
#             print("Generating the Sales in UC Report...")
#             generate_sales_in_uc_report("input_file.xlsx", writer)
#     except Exception as e:
#         print(f"An error occurred in the Sales in UC Report module: {e}")

import pandas as pd

def generate_sales_in_uc_report(input_file, writer):
    """
    Generates the Sales in UC report (Pack Wise and Brand Wise) and writes them to the output Excel file.

    Args:
        input_file (str): Path to the input Excel file.
        writer (pd.ExcelWriter): Excel writer object to write the output file.
    """
    try:
        print("Loading movements and products sheets...")
        movements = pd.read_excel(input_file, sheet_name="movements")
        products = pd.read_excel(input_file, sheet_name="products")
    except ValueError as e:
        print("Error loading the sheets. Ensure the input file contains 'movements' and 'products'.")
        raise ValueError("The input file must contain 'movements' and 'products' sheets.") from e
    except Exception as e:
        print(f"An error occurred while loading the sheets: {e}")
        raise

    # Merge movements with products on "Code" (SKU Code)
    try:
        print("Merging movements with products on 'Code' (SKU Code)...")
        merged_data = pd.merge(
            movements,
            products[["Code", "Pack", "Brand", "Unitfactor"]],
            how="left",
            left_on="Code",
            right_on="Code",
        )
    except Exception as e:
        print(f"An error occurred while merging data: {e}")
        raise

    # Warn if there are unmatched product codes
    unmatched_codes = merged_data[merged_data["Pack"].isna()]
    if not unmatched_codes.empty:
        print(f"Warning: {len(unmatched_codes)} product codes in 'movements' were not found in 'products'.")

    # Adjust quantity by multiplying with the Unitfactor
    merged_data["AdjustedQuantity"] = merged_data["Quantity"] * merged_data["Unitfactor"]

    # Generate the Pack Wise Sale in UC report
    try:
        print("Creating the pivot table for Pack Wise Sale in UC Report...")
        pack_pivot = merged_data.pivot_table(
            index="RouteCode",
            columns="Pack",
            values="AdjustedQuantity",
            aggfunc="sum",
            fill_value=0,
        )
        # Round to two decimal places
        pack_pivot = pack_pivot.round(2)

        # Add a "Total" row at the end of the pack pivot table
        pack_total_row = pack_pivot.sum(numeric_only=True, skipna=True).round(2)
        pack_total_row.name = "Total"
        pack_pivot = pd.concat([pack_pivot, pack_total_row.to_frame().T])
    except Exception as e:
        print(f"An error occurred while creating the pack pivot table: {e}")
        raise

    # Reset the index to make it tabular
    pack_wise_report = pack_pivot.reset_index()

    # Write the Pack Wise Sale in UC report to the Excel file
    try:
        print("Writing the Pack Wise Sale in UC Report to the Excel sheet...")
        pack_wise_report.to_excel(writer, sheet_name="Pack Wise Sale in UC", index=False, startrow=1)
    except Exception as e:
        print(f"An error occurred while writing the report to the Excel file: {e}")
        raise

    # Format the Pack Wise Sale in UC sheet
    try:
        print("Formatting the Pack Wise Sale in UC Excel sheet...")
        workbook = writer.book
        worksheet = writer.sheets["Pack Wise Sale in UC"]

        worksheet.write(0, 0, "Pack Wise Sale in UC Report")

        header_format = workbook.add_format({
            "bold": True,
            "text_wrap": True,
            "valign": "center",
            "align": "center",
            "fg_color": "#D7E4BC",
            "border": 1,
        })

        for col_num, value in enumerate(pack_wise_report.columns.values):
            worksheet.write(1, col_num, value, header_format)

        for col_num, col_name in enumerate(pack_wise_report.columns):
            max_width = max(pack_wise_report[col_name].astype(str).apply(len).max(), len(str(col_name))) + 2
            worksheet.set_column(col_num, col_num, max_width)
        
        # Apply number formatting to numeric columns
        number_format = workbook.add_format({'num_format': '0.00'})
        for col_num in range(1, len(pack_wise_report.columns)):  # Start from 1 to skip 'RouteCode'
            worksheet.set_column(col_num, col_num, None, number_format)
    except Exception as e:
        print(f"An error occurred while formatting the Excel sheet: {e}")
        raise

    # Generate the Brand Wise Sale in UC report
    try:
        print("Creating the pivot table for Brand Wise Sale in UC Report...")
        brand_pivot = merged_data.pivot_table(
            index="RouteCode",
            columns="Brand",
            values="AdjustedQuantity",
            aggfunc="sum",
            fill_value=0,
        )
        # Round to two decimal places
        brand_pivot = brand_pivot.round(2)

        # Add a "Total" row at the end of the brand pivot table
        brand_total_row = brand_pivot.sum(numeric_only=True, skipna=True).round(2)
        brand_total_row.name = "Total"
        brand_pivot = pd.concat([brand_pivot, brand_total_row.to_frame().T])
    except Exception as e:
        print(f"An error occurred while creating the brand pivot table: {e}")
        raise

    # Reset the index to make it tabular
    brand_wise_report = brand_pivot.reset_index()

    # Write the Brand Wise Sale in UC report to the Excel file
    try:
        print("Writing the Brand Wise Sale in UC Report to the Excel sheet...")
        brand_wise_report.to_excel(writer, sheet_name="Brand Wise Sale in UC", index=False, startrow=1)
    except Exception as e:
        print(f"An error occurred while writing the report to the Excel file: {e}")
        raise

    # Format the Brand Wise Sale in UC sheet
    try:
        print("Formatting the Brand Wise Sale in UC Excel sheet...")
        worksheet = writer.sheets["Brand Wise Sale in UC"]

        worksheet.write(0, 0, "Brand Wise Sale in UC Report")

        for col_num, value in enumerate(brand_wise_report.columns.values):
            worksheet.write(1, col_num, value, header_format)

        for col_num, col_name in enumerate(brand_wise_report.columns):
            max_width = max(brand_wise_report[col_name].astype(str).apply(len).max(), len(str(col_name))) + 2
            worksheet.set_column(col_num, col_num, max_width)
        
        # Apply number formatting to numeric columns
        for col_num in range(1, len(brand_wise_report.columns)):  # Start from 1 to skip 'RouteCode'
            worksheet.set_column(col_num, col_num, None, number_format)
    except Exception as e:
        print(f"An error occurred while formatting the Excel sheet: {e}")
        raise

    print("Sales in UC Report generation complete.")

if __name__ == "__main__":
    try:
        with pd.ExcelWriter("output_file.xlsx", engine="xlsxwriter") as writer:
            print("Generating the Sales in UC Report...")
            generate_sales_in_uc_report("input_file.xlsx", writer)
    except Exception as e:
        print(f"An error occurred in the Sales in UC Report module: {e}")
