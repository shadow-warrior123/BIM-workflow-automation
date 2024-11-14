import pandas as pd
import os

def find_level_column(df, params):
    for param in params:
        if param in df.columns:
            return param
    return None

def sanitize_sheet_name(name):
    """Ensure the sheet name is <= 31 characters."""
    if len(name) > 31:
        return name[:28] + '...'
    return name

def reformat_and_save_excel(input_file_path, output_file_path, level_params):
    workbook = pd.ExcelFile(input_file_path)
    sheets_data = []

    for sheet_name in workbook.sheet_names:
        df = pd.read_excel(workbook, sheet_name=sheet_name, header=0)
        sheets_data.append((sheet_name, df))

    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
        for sheet_name, df in sheets_data:
            level_col = find_level_column(df, level_params)
            if level_col is not None:
                grouped = df.groupby(level_col)
                for level, group in grouped:
                    sanitized_sheet_name = sanitize_sheet_name(f'{sheet_name}_{level}')
                    group.to_excel(writer, sheet_name=sanitized_sheet_name, index=False)
            else:
                sanitized_sheet_name = sanitize_sheet_name(sheet_name)
                df.to_excel(writer, sheet_name=sanitized_sheet_name, index=False)

    print(f"Reformatted data saved to: {output_file_path}")

if __name__ == "__main__":
    input_files = [
        "C:\\Users\\f2021\\Desktop\\output\\Cleaned_Sorted_Area_Data.xlsx",
        "C:\\Users\\f2021\\Desktop\\output\\Cleaned_Sorted_Volume_Data.xlsx"
    ]

    output_folder = "C:\\Users\\f2021\\Desktop\\output"
    output_files = [
        os.path.join(output_folder, "Reformatted_Area_Data.xlsx"),
        os.path.join(output_folder, "Reformatted_Volume_Data.xlsx")
    ]

    level_params = ['Level', 'Base Level', 'Base Constraint', 'Work plane']

    for input_file, output_file in zip(input_files, output_files):
        reformat_and_save_excel(input_file, output_file, level_params)
