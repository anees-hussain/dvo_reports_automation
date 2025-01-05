import pandas as pd

def generate_avg_drop_size_report(input_file, writer):
    """
    Generates the Average Drop Size of Invoices Route Wise report and writes it to the output Excel file.
    """
    try:
        print("Loading movements sheet...")
        movements = pd.read_excel(input_file, sheet_name="movements")
    except ValueError as e:
        print("Error loading the movements sheet. Ensure the input file contains 'movements'.")
        raise ValueError("The input file must contain 'movements' sheet.") from e
    except Exception as e:
        print(f"An error occurred while loading the sheet: {e}")
        raise

    # Step to calculate Average Drop Size of Invoices Route Wise
    try:
        print("Calculating Average Drop Size of Invoices Route Wise...")

        # Get unique SlipNr and sum quantities
        unique_slipnr = movements.groupby("SlipNr")["Quantity"].sum().reset_index()

        # Merge with RouteCode based on SlipNr
        merged_slipnr = pd.merge(unique_slipnr, movements[["SlipNr", "RouteCode"]], on="SlipNr", how="left").drop_duplicates()

        # Save the unique slipNr with quantities
        unique_slipnr_file = "unique_slipnr_with_quantities.xlsx"
        with pd.ExcelWriter(unique_slipnr_file, engine="xlsxwriter") as slipnr_writer:
            merged_slipnr.to_excel(slipnr_writer, sheet_name="unique slipNr with quantities", index=False, startrow=1)

        # Calculate the average drop size per RouteCode
        avg_drop_size = merged_slipnr.groupby("RouteCode")["Quantity"].agg(["sum", "count"]).reset_index()
        avg_drop_size["Avg Drop Size"] = round(avg_drop_size["sum"] / avg_drop_size["count"],2)

        # Write the Average Drop Size of Invoices Route Wise to the new sheet
        avg_drop_size.to_excel(writer, sheet_name="Avg DropSize", index=False, startrow=1)

        # Formatting the new sheet
        workbook = writer.book
        worksheet = writer.sheets["Avg DropSize"]
        worksheet.write(0, 0, "Avg DropSize")

        header_format = workbook.add_format({
            "bold": True,
            "text_wrap": True,
            "valign": "center",
            "align": "center",
            "fg_color": "#D7E4BC",
            "border": 1,
        })

        for col_num, value in enumerate(avg_drop_size.columns.values):
            worksheet.write(1, col_num, value, header_format)

        for col_num, col_name in enumerate(avg_drop_size.columns):
            max_width = max(avg_drop_size[col_name].astype(str).apply(len).max(), len(str(col_name))) + 2
            worksheet.set_column(col_num, col_num, max_width)

    except Exception as e:
        print(f"An error occurred while calculating Average Drop Size of Invoices Route Wise: {e}")
        raise

    print("Average Drop Size of Invoices Route Wise Report generation complete.")


if __name__ == "__main__":
    try:
        with pd.ExcelWriter("output_file.xlsx", engine="xlsxwriter") as writer:
            print("Generating the Average Drop Size of Invoices Route Wise Report...")
            generate_avg_drop_size_report("input_file.xlsx", writer)
    except Exception as e:
        print(f"An error occurred in the Average Drop Size Report module: {e}")
