import pandas as pd

def generate_avg_sku_per_invoice_report(input_file, writer):
    """
    Generates the Average SKU per Invoice Report and writes it to the output Excel file.

    Args:
        input_file (str): Path to the input Excel file.
        writer (pd.ExcelWriter): Excel writer object to write the output file.
    """
    try:
        print("Loading movements sheet...")
        # Load movements sheet
        movements = pd.read_excel(input_file, sheet_name="movements")
        
        # Validate required columns
        print("Validating required columns...")
        required_columns = ["RouteCode", "Code", "SlipNr", "Date"]
        if not all(col in movements.columns for col in required_columns):
            raise ValueError(f"The input data is missing required columns: {', '.join(required_columns)}")

        # Calculate total SKUs and unique invoices for each route code
        print("Calculating total SKUs and unique invoices for each route code...")
        overall_summary = movements.groupby("RouteCode").agg(
            total_skus=("Code", "count"),
            unique_invoices=("SlipNr", "nunique")
        ).reset_index()

        # Calculate SKU/Invoice with division-by-zero handling
        print("Calculating SKU/Invoice for each route code...")
        overall_summary["SKU/Invoice"] = overall_summary.apply(
            lambda row: round(row["total_skus"] / row["unique_invoices"],2) if row["unique_invoices"] > 0 else 0,
            axis=1
        )
        overall_summary.columns = ["Route Code", "Number of SKUs", "Number of Unique Invoices", "SKU/Invoice"]

        # Get last date from the dataset
        print("Getting the last date from the dataset...")

        # Normalize 'Date' column to remove time portion
        movements["Date"] = pd.to_datetime(movements["Date"]).dt.date

        # Get the most recent date
        last_date = movements["Date"].max()

        # Subtract 1 day to get the previous day
        previous_day = last_date - pd.Timedelta(days=1)

        # Filter data for the previous day
        print(f"Filtering data for the previous day: {previous_day}...")
        last_day_data = movements[movements["Date"] == previous_day]

        # Calculate summary for the previous day
        print("Calculating summary for the previous day...")
        last_day_summary = last_day_data.groupby("RouteCode").agg(
            total_skus=("Code", "count"),
            unique_invoices=("SlipNr", "nunique")
        ).reset_index()

        # Calculate SKU/Invoice for the previous day
        print("Calculating SKU/Invoice for the previous day...")
        last_day_summary["SKU/Invoice"] = last_day_summary.apply(
            lambda row: round(row["total_skus"] / row["unique_invoices"],2) if row["unique_invoices"] > 0 else 0,
            axis=1
        )
        last_day_summary.columns = ["Route Code", "Number of SKUs", "Number of Unique Invoices", "SKU/Invoice"]

        # Write both tables to the sheet
        print("Writing summary data to the Excel sheet...")
        overall_start_row = 1
        last_day_start_row = len(overall_summary) + 5
        overall_summary.to_excel(writer, sheet_name="Avg SKU per Invoice", index=False, startrow=overall_start_row)
        last_day_summary.to_excel(writer, sheet_name="Avg SKU per Invoice", index=False, startrow=last_day_start_row)

        # Format the sheet
        print("Formatting the Excel sheet...")
        workbook = writer.book
        worksheet = writer.sheets["Avg SKU per Invoice"]

        # Add headers for each table
        worksheet.write(0, 0, "Overall SKU/Invoice Report")
        worksheet.write(last_day_start_row - 1, 0, f"Previous Day SKU/Invoice Report ({previous_day.strftime('%Y-%m-%d')})")

        # Add formatting
        header_format = workbook.add_format({
            "bold": True,
            "text_wrap": True,
            "valign": "center",
            "align": "center",
            "fg_color": "#F4B084",
            "border": 1,
        })

        # Apply formatting to headers
        for col_num, value in enumerate(overall_summary.columns.values):
            worksheet.write(overall_start_row, col_num, value, header_format)

        for col_num, value in enumerate(last_day_summary.columns.values):
            worksheet.write(last_day_start_row, col_num, value, header_format)

        # Adjust column widths
        print("Adjusting column widths...")
        column_widths = [max(len(str(val)), 15) for val in overall_summary.columns]
        for i, width in enumerate(column_widths):
            worksheet.set_column(i, i, width)

        print("Report generation complete.")

    except ValueError as ve:
        print(f"Validation error: {ve}")
        raise RuntimeError(f"An error occurred while generating the SKU/Invoice report: {ve}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        raise RuntimeError(f"An error occurred while generating the SKU/Invoice report: {e}")
