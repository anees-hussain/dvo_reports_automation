import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd

class ReportGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DVO Reports Automation")
        self.root.geometry("600x300")
        self.root.configure(bg="white")
        self.setup_ui()

    def setup_ui(self):
        # Title Label
        title_label = tk.Label(
            self.root,
            text="DVO Reports Automation",
            font=("Helvetica", 16, "bold"),
            bg="white",
        )
        title_label.pack(pady=20)

        # Upload Button
        upload_button = tk.Button(
            self.root,
            text="Upload File",
            font=("Helvetica", 14),
            bg="#4CAF50",
            fg="white",
            padx=20,
            pady=10,
            command=self.upload_file,
        )
        upload_button.pack(pady=20)

        # Footer
        footer_label = tk.Label(
            self.root,
            text="Developed by CH M Anees Hussain Toor (aneeshussain009@gmail.com)",
            font=("Helvetica", 10),
            bg="white",
        )

        footer_label.pack(side="bottom", pady=10)

    def upload_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xls;*.xlsx")]
        )
        if file_path:
            try:
                print(f"File selected: {file_path}")  # Debugging
                self.validate_file(file_path)
                messagebox.showinfo("Success", "File processed successfully!")
            except ValueError as e:
                print(f"Validation Error: {str(e)}")  # Debugging
                messagebox.showerror("Validation Error", str(e))
            except Exception as e:
                print(f"Unexpected Error: {str(e)}")  # Debugging
                messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")

    def validate_file(self, file_path):
        required_sheets = ["movements", "outlets", "products"]
        required_columns = {
            "movements": [
                "Code", "Name", "Date", "Quantity", "SlipNr", "OutletCode",
                "OutletName", "RouteCode"
            ],
            "outlets": ["OutletCode", "OutletName", "RouteCode"],
            "products": ["Code", "SKUDescription", "Pack", "Brand", "Unitfactor"],
        }

        # Load the Excel file
        try:
            print(f"Loading Excel file: {file_path}")  # Debugging
            excel_data = pd.ExcelFile(file_path)
        except Exception as e:
            print(f"Error loading Excel file: {str(e)}")  # Debugging
            raise ValueError(f"Unable to load file '{file_path}': {e}")

        # Validate sheets
        missing_sheets = [sheet for sheet in required_sheets if sheet not in excel_data.sheet_names]
        if missing_sheets:
            print(f"Missing sheets: {missing_sheets}")  # Debugging
            raise ValueError(f"Missing sheets: {', '.join(missing_sheets)}")

        # Validate columns
        for sheet, columns in required_columns.items():
            try:
                print(f"Validating sheet: {sheet}")  # Debugging
                sheet_data = pd.read_excel(file_path, sheet_name=sheet)
                normalized_columns = sheet_data.columns.str.strip()
                missing_columns = [col for col in columns if col not in normalized_columns]

                if missing_columns:
                    print(f"Missing columns in {sheet}: {missing_columns}")  # Debugging
                    raise ValueError(
                        f"Missing columns in sheet '{sheet}': {', '.join(missing_columns)}"
                    )
            except Exception as e:
                print(f"Error reading sheet '{sheet}': {str(e)}")  # Debugging
                raise ValueError(f"Error reading sheet '{sheet}': {e}")

        # If validation passes, generate reports
        self.generate_reports(file_path)

    def generate_reports(self, file_path):
        output_dir = filedialog.askdirectory(title="Select Output Directory")
        if not output_dir:
            print("No output directory selected.")  # Debugging
            raise ValueError("Output directory not selected.")

        output_file = os.path.join(output_dir, "Generated_Reports.xlsx")
        try:
            print(f"Generating reports...")  # Debugging
            from modules.pack_wise_report import generate_pack_wise_report
            from modules.brand_wise_report import generate_brand_wise_report
            from modules.availability_report import generate_availability_report
            from modules.penetration_report import generate_penetration_report
            from modules.avg_sku_per_invoice import generate_avg_sku_per_invoice_report
            from modules.sales_in_uc import generate_sales_in_uc_report
            from modules.avg_drop_size import generate_avg_drop_size_report

            with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
                generate_pack_wise_report(file_path, writer)
                generate_brand_wise_report(file_path, writer)
                generate_sales_in_uc_report(file_path, writer)
                generate_availability_report(file_path, writer)
                generate_penetration_report(file_path, writer)
                generate_avg_sku_per_invoice_report(file_path, writer)
                generate_avg_drop_size_report(file_path, writer)

            messagebox.showinfo(
                "Success", f"Reports generated successfully! Saved at {output_file}"
            )
            print(f"Reports saved at: {output_file}")  # Debugging
        except Exception as e:
            print(f"Error generating reports in main.py : {str(e)}")  # Debugging
            raise Exception(f"Error generating reports in main.py : {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ReportGeneratorApp(root)
    root.mainloop()