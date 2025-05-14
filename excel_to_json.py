import os
import pandas as pd
import json
from tkinter import Tk
from tkinter.filedialog import askopenfilename, askdirectory

def excel_to_json(excel_file, json_file):
    """
    Converts an Excel file back to the JSON format required for GST portal.

    :param excel_file: Path to the input Excel file.
    :param json_file: Path to the output JSON file.
    """
    try:
        # Read the Excel file
        b2b_df = pd.read_excel(excel_file, sheet_name='B2B')
        b2cs_df = pd.read_excel(excel_file, sheet_name='B2CS')
        hsn_df = pd.read_excel(excel_file, sheet_name='HSN')
        doc_issue_df = pd.read_excel(excel_file, sheet_name='Doc Issue')

        # Ensure 'uqc' column in HSN data is treated as text and replace NA with "NA"
        if 'uqc' in hsn_df.columns:
            hsn_df['uqc'] = hsn_df['uqc'].fillna("NA").astype(str)

        # Convert B2B data back to JSON structure
        b2b_data = []
        for ctin, group in b2b_df.groupby('CTIN'):
            inv_list = []
            for _, row in group.iterrows():
                inv = {
                    "inum": row["Invoice Number"],
                    "idt": row["Invoice Date"],
                    "val": row["Invoice Value"],
                    "pos": row["Place of Supply"],
                    "rchrg": row["Reverse Charge"],
                    "inv_typ": row["Invoice Type"],
                    "itms": [
                        {
                            "num": row["Item Number"],
                            "itm_det": {
                                "txval": row["Taxable Value"],
                                "rt": row["Rate"],
                                "camt": row["Central Tax Amount"],
                                "samt": row["State Tax Amount"],
                                "csamt": row["Cess Amount"]
                            }
                        }
                    ]
                }
                inv_list.append(inv)
            b2b_data.append({"ctin": ctin, "inv": inv_list})

        # Convert B2CS data back to JSON structure
        b2cs_data = b2cs_df.to_dict(orient='records')

        # Convert HSN data back to JSON structure
        hsn_data = {"data": hsn_df.to_dict(orient='records')}

        # Convert Doc Issue data back to JSON structure
        doc_issue_data = []
        for doc_type, group in doc_issue_df.groupby('Document Type'):
            docs = []
            for _, row in group.iterrows():
                docs.append({
                    "from": row["From"],
                    "to": row["To"],
                    "totnum": row["Total Number"],
                    "cancel": row["Cancelled"],
                    "net_issue": row["Net Issued"]
                })
            doc_issue_data.append({"doc_typ": doc_type, "docs": docs})

        # Ask the user for GSTIN and filing period (fp)
        gstin = input("Enter GSTIN: ").strip()
        fp = input("Enter Filing Period (e.g., 022024): ").strip()

        # Combine all data into the final JSON structure
        final_json = {
            "gstin": gstin,
            "fp": fp,
            "gt": 0.00,  # Replace with the actual gross turnover if needed
            "cur_gt": 0.00,  # Replace with the actual current gross turnover if needed
            "b2b": b2b_data,
            "b2cs": b2cs_data,
            "hsn": hsn_data,
            "doc_issue": {"doc_det": doc_issue_data}
        }

        # Write the final JSON to the output file
        with open(json_file, 'w') as file:
            json.dump(final_json, file, indent=4)

        print(f"Successfully converted {excel_file} back to {json_file}")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    # Hide the root Tkinter window
    Tk().withdraw()

    # Ask the user to select the original JSON file
    original_json = askopenfilename(
        title="Select Original JSON File",
        filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
    )
    if not original_json:
        print("No JSON file selected. Exiting.")
        exit()

    # Extract the name of the original JSON file (without extension)
    original_json_name = os.path.basename(original_json)

    # Ask the user to select the Excel file
    input_excel = askopenfilename(
        title="Select Edited Excel File",
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
    )
    if not input_excel:
        print("No Excel file selected. Exiting.")
        exit()

    # Ask the user to select the folder to save the new JSON file
    output_folder = askdirectory(title="Select Folder to Save JSON File")
    if not output_folder:
        print("No folder selected. Exiting.")
        exit()

    # Construct the full path for the new JSON file
    output_json = os.path.join(output_folder, original_json_name)

    # Check if a file with the same name exists and replace it
    if os.path.exists(output_json):
        print(f"File {output_json} already exists. It will be replaced.")

    # Convert Excel to JSON
    excel_to_json(input_excel, output_json)

    print(f"New JSON file saved as: {output_json}")
