import os
import pandas as pd
import json
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename

def combine_json_to_excel(json_files, excel_file):
    """
    Combines multiple JSON files with the same structure into a single Excel file.

    :param json_files: List of paths to the input JSON files.
    :param excel_file: Path to the output Excel file.
    """
    try:
        # Initialize empty lists for each type of data
        b2b_data = []
        b2cs_data = []
        hsn_data = []
        doc_issue_data = []

        # Process each JSON file
        for json_file in json_files:
            with open(json_file, 'r') as file:
                data = json.load(file)

            # Extract and flatten 'b2b' data
            for entry in data.get('b2b', []):
                ctin = entry.get('ctin')
                for invoice in entry.get('inv', []):
                    invoice_data = {
                        'CTIN': ctin,
                        'Invoice Number': invoice.get('inum'),
                        'Invoice Date': invoice.get('idt'),
                        'Invoice Value': invoice.get('val'),
                        'Place of Supply': invoice.get('pos'),
                        'Reverse Charge': invoice.get('rchrg'),
                        'Invoice Type': invoice.get('inv_typ')
                    }
                    for item in invoice.get('itms', []):
                        item_data = {
                            'Item Number': item.get('num'),
                            'Taxable Value': item['itm_det'].get('txval'),
                            'Rate': item['itm_det'].get('rt'),
                            'Central Tax Amount': item['itm_det'].get('camt'),
                            'State Tax Amount': item['itm_det'].get('samt'),
                            'Cess Amount': item['itm_det'].get('csamt')
                        }
                        b2b_data.append({**invoice_data, **item_data})

            # Extract 'b2cs' data
            b2cs_data.extend(data.get('b2cs', []))

            # Extract 'hsn' data
            hsn_data.extend(data.get('hsn', {}).get('data', []))

            # Extract 'doc_issue' data
            for doc in data.get('doc_issue', {}).get('doc_det', []):
                doc_type = doc.get('doc_typ')
                for detail in doc.get('docs', []):
                    doc_issue_data.append({
                        'Document Type': doc_type,
                        'From': detail.get('from'),
                        'To': detail.get('to'),
                        'Total Number': detail.get('totnum'),
                        'Cancelled': detail.get('cancel'),
                        'Net Issued': detail.get('net_issue')
                    })

        # Convert data to DataFrames
        b2b_df = pd.DataFrame(b2b_data)
        b2cs_df = pd.DataFrame(b2cs_data)
        hsn_df = pd.DataFrame(hsn_data)
        doc_issue_df = pd.DataFrame(doc_issue_data)

        # Ensure 'uqc' column in HSN data is treated as text
        if not hsn_df.empty and 'uqc' in hsn_df.columns:
            hsn_df['uqc'] = hsn_df['uqc'].astype(str)

        # Group and sum up B2CS data if columns (rt, sply_ty, pos, typ) are repeating
        if not b2cs_df.empty:
            original_columns = b2cs_df.columns.tolist()  # Save original column order
            b2cs_df = b2cs_df.groupby(['rt', 'sply_ty', 'pos', 'typ'], as_index=False).sum()
            b2cs_df = b2cs_df[original_columns]  # Reorder columns to original order

        # Group and sum up HSN data if columns (hsn_sc, uqc, rt) are repeating
        if not hsn_df.empty:
            original_columns = hsn_df.columns.tolist()  # Save original column order
            hsn_df = hsn_df.groupby(['hsn_sc', 'uqc', 'rt'], as_index=False).sum()
            hsn_df['num'] = range(1, len(hsn_df) + 1)  # Update 'num' column with serial numbers
            hsn_df = hsn_df[original_columns]  # Reorder columns to original order

        # Write all DataFrames to an Excel file with multiple sheets
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            if not b2b_df.empty:
                b2b_df.to_excel(writer, sheet_name='B2B', index=False)
            if not b2cs_df.empty:
                b2cs_df.to_excel(writer, sheet_name='B2CS', index=False)
            if not hsn_df.empty:
                hsn_df.to_excel(writer, sheet_name='HSN', index=False)
            if not doc_issue_df.empty:
                doc_issue_df.to_excel(writer, sheet_name='Doc Issue', index=False)

        print(f"Successfully combined JSON files into {excel_file}")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    # Hide the root Tkinter window
    Tk().withdraw()

    # Initialize an empty list to store selected JSON file paths
    input_json_files = []

    # Allow the user to browse and select JSON files multiple times
    while True:
        json_file = askopenfilename(
            title="Select a JSON File",
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        if not json_file:
            break  # Stop if the user cancels the dialog
        input_json_files.append(json_file)
        print(f"Selected: {json_file}")

    if not input_json_files:
        print("No JSON files selected. Exiting.")
        exit()

    # Ask the user to provide a name and location for the Excel file
    output_excel = asksaveasfilename(
        title="Save Excel File As",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
    )
    if not output_excel:
        print("No file name provided. Exiting.")
        exit()

    # Combine JSON files into a single Excel file
    combine_json_to_excel(input_json_files, output_excel)