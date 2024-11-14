import pandas as pd
import os

# Updated list of possible floor levels to match
possible_floors = ['First Floor', 'Foundation Floor', 'Third Floor', 'Second Floor', 'Ground Floor', 'Terrace Floor']

# Function to extract category and level from sheet name
def extract_category_level(sheet_name):
    parts = sheet_name.split('_')
    if len(parts) > 1:
        category = '_'.join(parts[:-1])
        level = parts[-1]
    else:
        category = parts[0]
        level = ""

    # Match level with possible floors based on the most similar match
    if level:
        best_match = max(possible_floors, key=lambda x: sum(a == b for a, b in zip(x.lower(), level.lower())))
        if sum(a == b for a, b in zip(best_match.lower(), level.lower())) >= 3:  # At least 3 characters match
            level = best_match

    return category, level

# Function to process the Excel file
def process_excel_file(input_file, output_file_path):
    # Load the provided Excel file to get the detailed data without 'Element ID' column
    excel_data = pd.ExcelFile(input_file)

    # Initialize a dictionary to hold data categorized by levels
    levels_data_formatted = {}

    # Iterate through each sheet and categorize the data by cleaned levels
    for sheet in excel_data.sheet_names:
        sheet_data = pd.read_excel(input_file, sheet_name=sheet)
        sheet_data = sheet_data.rename(columns={sheet_data.columns[0]: 'Material', sheet_data.columns[1]: 'Total Quantity Used'})
        category, level = extract_category_level(sheet)
        if level not in levels_data_formatted:
            levels_data_formatted[level] = []
        levels_data_formatted[level].append((category, sheet_data))

    # Create a new Excel writer object for the updated file
    with pd.ExcelWriter(output_file_path) as writer:
        for level, data_list in levels_data_formatted.items():
            combined_data = pd.DataFrame()
            for category, data in data_list:
                # Add category name as a header
                category_header = pd.DataFrame([[category, '']], columns=['Material', 'Total Quantity Used'])
                # Add table header for materials table
                combined_data = pd.concat([combined_data, category_header], ignore_index=True)
                combined_data = pd.concat([combined_data, data], ignore_index=True)
                # Add three rows of gap
                combined_data = pd.concat([combined_data, pd.DataFrame([['', '']] * 3, columns=['Material', 'Total Quantity Used'])], ignore_index=True)
            combined_data.to_excel(writer, sheet_name=level, index=False)

    print(f"Processed file saved to {output_file_path}")

# Example usage
if __name__ == "__main__":
    input_files = [
        "C:\\Users\\f2021\\Desktop\\output\\Processed_Area_Data.xlsx",
        "C:\\Users\\f2021\\Desktop\\output\\Processed_Volume_Data.xlsx"
    ]

    output_folder = "C:\\Users\\f2021\\Desktop\\Rajiv_Sir_Project\\final_output"

    output_files = [
        os.path.join(output_folder, "Final_Area_Data.xlsx"),
        os.path.join(output_folder, "Final_Volume_Data.xlsx")
    ]

    for input_file, output_file in zip(input_files, output_files):
        process_excel_file(input_file, output_file)
