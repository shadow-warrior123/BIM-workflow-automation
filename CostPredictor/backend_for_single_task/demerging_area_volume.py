import pandas as pd
import os

def process_data(df):
    # Ensure column names are correctly identified
    area_columns = [col for col in df.columns if col.lower().endswith('area')]
    volume_columns = [col for col in df.columns if col.lower().endswith('volume')]

    # Include Category column at first position
    area_data = df[['Category', 'Element ID', 'Family', 'Type', 'Mark', 'Level'] + area_columns]
    volume_data = df[['Category', 'Element ID', 'Family', 'Type', 'Mark', 'Level'] + volume_columns]

    # Remove rows with all NaN values in area columns
    area_data = area_data.dropna(subset=area_columns, how='all')

    # Remove rows with all NaN values in volume columns
    volume_data = volume_data.dropna(subset=volume_columns, how='all')

    # Remove completely empty columns (only header)
    area_data = area_data.dropna(axis=1, how='all')
    volume_data = volume_data.dropna(axis=1, how='all')

    save_excel(area_data, r"C:\Users\f2021\Desktop\output\area_data.xlsx")
    save_excel(volume_data, r"C:\Users\f2021\Desktop\output\volume_data.xlsx")

def save_excel(df, file_path):
    try:
        df.to_excel(file_path, index=False)
        print(f"File saved successfully: {file_path}")
    except PermissionError:
        print(f"Permission denied: {file_path}. Please close the file if it's open and try again.")
    except Exception as e:
        print(f"Failed to save the Excel file: {e}")

def main():
    input_file = r"C:\Users\f2021\Desktop\output\Final_Merged_Sorted_Excel_Sheet_with_Level.xlsx"  # Replace with your input file path
    try:
        df = pd.read_excel(input_file)
        process_data(df)
    except Exception as e:
        print(f"Failed to read the Excel file: {e}")

if __name__ == "__main__":
    main()
