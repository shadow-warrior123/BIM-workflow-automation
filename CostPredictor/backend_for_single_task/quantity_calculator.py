import pandas as pd
import os

def calculate_material_totals(input_file, output_file_path):
    # Load the Excel file
    excel_data = pd.ExcelFile(input_file)
    sheet_names = excel_data.sheet_names

    # Initialize a dictionary to hold the data for each sheet
    sheets_data = {}

    # Iterate through each sheet and process the data
    for sheet in sheet_names:
        sheet_data = pd.read_excel(input_file, sheet_name=sheet)
        if 'Element ID' in sheet_data.columns:
            sheet_data = sheet_data.drop(columns=['Element ID'])
        material_totals = sheet_data.select_dtypes(include='number').sum()
        sheets_data[sheet] = material_totals

    # Create a new Excel writer object
    with pd.ExcelWriter(output_file_path) as writer:
        for sheet, data in sheets_data.items():
            # Convert the Series to DataFrame for better formatting
            material_totals_df = data.reset_index()
            material_totals_df.columns = ['Material', 'Total Quantity Used']
            # Write each sheet's data to the Excel file
            material_totals_df.to_excel(writer, sheet_name=sheet, index=False)

    print(f"Processed file saved to {output_file_path}")

# Example usage
if __name__ == "__main__":
    input_files = [
        "C:\\Users\\f2021\\Desktop\\output\\Reformatted_Area_Data.xlsx",
        "C:\\Users\\f2021\\Desktop\\output\\Reformatted_Volume_Data.xlsx"
    ]

    output_files = [
        "C:\\Users\\f2021\\Desktop\\output\\Processed_Area_Data.xlsx",
        "C:\\Users\\f2021\\Desktop\\output\\Processed_Volume_Data.xlsx"
    ]

    for input_file, output_file in zip(input_files, output_files):
        calculate_material_totals(input_file, output_file)
