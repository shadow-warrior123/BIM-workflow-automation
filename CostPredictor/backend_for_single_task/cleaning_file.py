import pandas as pd
import os

# Function to remove rows where Total Quantity Used is 0 for a given sheet
def clean_sheet(sheet_name, excel_data):
    df = pd.read_excel(excel_data, sheet_name=sheet_name)
    cleaned_df = df[df['Total Quantity Used'] != 0]
    return cleaned_df

# Function to process an Excel file
def process_file(input_file_path, output_folder, prefix="Presentable"):
    # Load the Excel file
    excel_data = pd.ExcelFile(input_file_path)

    # Get all sheet names
    sheet_names = excel_data.sheet_names

    # Clean all sheets and save them back to a new Excel file
    cleaned_sheets = {}
    for sheet in sheet_names:
        cleaned_sheets[sheet] = clean_sheet(sheet, excel_data)

    # Create the output file name
    input_file_name = os.path.basename(input_file_path)
    output_file_name = f"{prefix}_{input_file_name}"
    output_file_path = os.path.join(output_folder, output_file_name)

    # Save the cleaned data to a new Excel file
    with pd.ExcelWriter(output_file_path) as writer:
        for sheet_name, cleaned_df in cleaned_sheets.items():
            cleaned_df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f'Cleaned file saved to {output_file_path}')

# Output folder
output_folder = r"C:\Users\f2021\Desktop\Rajiv_Sir_Project\final_output"

# Paths for the first file
input_file_path_1 = r"C:\Users\f2021\Desktop\Rajiv_Sir_Project\final_output\Final_Area_Data.xlsx"

# Paths for the second file
input_file_path_2 = r"C:\Users\f2021\Desktop\Rajiv_Sir_Project\final_output\Final_Volume_Data.xlsx"

# Process the first file
process_file(input_file_path_1, output_folder)

# Process the second file
process_file(input_file_path_2, output_folder)
