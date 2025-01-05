import pandas as pd

def generate_pack_wise_report(input_file, writer):
    """
    Generates the Pack Wise Report and writes it to the output Excel file.

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
            products[["Code", "Pack"]],
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

    # Pivot the data to get the Pack Wise Report
    try:
        print("Creating the pivot table for the Pack Wise Report...")
        pivot_table = merged_data.pivot_table(
            index="RouteCode",
            columns="Pack",
            values="Quantity",
            aggfunc="sum",
            fill_value=0,
        )
    except Exception as e:
        print(f"An error occurred while creating the pivot table: {e}")
        raise

    # Add a "Total" row at the end of the pivot table
    total_row = pivot_table.sum(numeric_only=True, skipna=True)
    total_row.name = "Total"  # Label the total row

    # Use pd.concat to append the total row
    pivot_table = pd.concat([pivot_table, total_row.to_frame().T])

    # Reset the index to make it tabular
    pack_wise_report = pivot_table.reset_index()

    # Write the report to the Excel file
    try:
        print("Writing the Pack Wise Report to the Excel sheet...")
        pack_wise_report.to_excel(writer, sheet_name="Pack Wise", index=False, startrow=1)
    except Exception as e:
        print(f"An error occurred while writing the report to the Excel file: {e}")
        raise

    # Format the sheet
    try:
        print("Formatting the Excel sheet...")
        workbook = writer.book
        worksheet = writer.sheets["Pack Wise"]

        # Add a descriptive title
        worksheet.write(0, 0, "Pack Wise Report")

        # Add a header format
        header_format = workbook.add_format({
            "bold": True,
            "text_wrap": True,
            "valign": "center",
            "align": "center",
            "fg_color": "#D7E4BC",
            "border": 1,
        })

        # Apply the header format to the first row
        for col_num, value in enumerate(pack_wise_report.columns.values):
            worksheet.write(1, col_num, value, header_format)

        # Adjust column widths dynamically
        for col_num, col_name in enumerate(pack_wise_report.columns):
            max_width = max(pack_wise_report[col_name].astype(str).apply(len).max(), len(str(col_name))) + 2
            worksheet.set_column(col_num, col_num, max_width)
    except Exception as e:
        print(f"An error occurred while formatting the Excel sheet: {e}")
        raise

    print("Pack Wise Report generation complete.")

if __name__ == "__main__":
    try:
        with pd.ExcelWriter("output_file.xlsx", engine="xlsxwriter") as writer:
            print("Generating the Pack Wise Report...")
            generate_pack_wise_report("input_file.xlsx", writer)
    except Exception as e:
        print(f"An error occurred in pack wise module: {e}")
