import pandas as pd
from pandas.errors import EmptyDataError, ParserError


def generate_availability_report(input_file, writer):
    """
    Generates the Availability Report and writes it to the output Excel file.

    Args:
        input_file (str): Path to the input Excel file. Must contain the sheets:
            - movements: Required columns ["Code", "Name", "Quantity", "OutletCode"]
            - products: Required columns ["Code", "SKUDescription"]
            - outlets: Required columns ["OutletCode", "OutletName"]
        writer (pd.ExcelWriter): Excel writer object to write the output file.

    Raises:
        FileNotFoundError: If the input file does not exist.
        ValueError: If required sheets or columns are missing.
    """
    try:
        # Load necessary sheets
        required_sheets = ["movements", "products", "outlets"]
        excel_data = pd.ExcelFile(input_file)

        # Validate sheets
        for sheet in required_sheets:
            if sheet not in excel_data.sheet_names:
                raise ValueError(f"Missing required sheet: '{sheet}' in the input file.")

        # Load sheets
        print("Loading sheets from input file...")
        movements = pd.read_excel(input_file, sheet_name="movements")
        products = pd.read_excel(input_file, sheet_name="products")
        outlets = pd.read_excel(input_file, sheet_name="outlets")

        # Print first few rows to verify data
        print("Movements Data:")
        print(movements.head())
        print("Products Data:")
        print(products.head())
        print("Outlets Data:")
        print(outlets.head())

        # Validate required columns
        required_columns = {
            "movements": ["Code", "Name", "Quantity", "OutletCode"],
            "products": ["Code", "SKUDescription"],
            "outlets": ["OutletCode", "OutletName"],
        }
        for sheet_name, columns in required_columns.items():
            sheet_data = locals()[sheet_name]
            missing_columns = [col for col in columns if col not in sheet_data.columns]
            if missing_columns:
                raise ValueError(
                    f"Missing required columns in '{sheet_name}': {', '.join(missing_columns)}"
                )

        # Merge movements with products and outlets to get SKU and Outlet details
        print("Merging data...")
        merged_data = pd.merge(
            movements,
            products[["Code", "SKUDescription"]],
            how="left",
            left_on="Code",
            right_on="Code",
        )
        merged_data = pd.merge(
            merged_data,
            outlets[["OutletCode", "OutletName"]],
            how="left",
            left_on="OutletCode",
            right_on="OutletCode",
        )

        # Print merged data to check for issues
        print("Merged Data:")
        print(merged_data.head())

        # Create a pivot table for the availability report
        print("Generating pivot table...")
        pivot_table = merged_data.pivot_table(
            index=["OutletCode", "OutletName_y"],
            columns="Name",
            values="Quantity",
            aggfunc="sum",
            fill_value=0,
        )

        # Reset the index to make it tabular
        availability_report = pivot_table.reset_index()

        # Write the report to the Excel file
        print("Writing report to Excel...")
        availability_report.to_excel(
            writer,
            sheet_name="Availability",
            index=False,
            startrow=1,
        )

        # Format the sheet
        workbook = writer.book
        worksheet = writer.sheets["Availability"]

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
        for col_num, value in enumerate(availability_report.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Adjust column widths dynamically
        for col_num, col_name in enumerate(availability_report.columns):
            column_width = max(
                len(str(col_name)),
                availability_report[col_name].astype(str).map(len).max()
            )
            worksheet.set_column(col_num, col_num, column_width)

    except FileNotFoundError:
        raise FileNotFoundError(f"The input file '{input_file}' was not found.")
    except EmptyDataError:
        raise ValueError("The input file is empty or corrupt.")
    except ParserError as e:
        raise ValueError(f"Error parsing the input file: {str(e)}")
    except Exception as e:
        print("Error in availability module:")
        print(f"Error Details: {str(e)}")
        raise ValueError(f"An unexpected error occurred in availability module: {str(e)}")


if __name__ == "__main__":
    try:
        writer = pd.ExcelWriter("output_file.xlsx", engine="xlsxwriter")
        generate_availability_report("input_file.xlsx", writer)
        writer.close()
        print("Availability report generated successfully!")
    except Exception as e:
        print(f"Error in availability report: {e}")
