import os
from tkinter import messagebox
import pandas as pd
import customtkinter
from helpers import get_current_fg_color

# Paths for the input files
input_file_paths = [
    "C:\\Users\\f2021\\Desktop\\output\\area_data.xlsx",
    "C:\\Users\\f2021\\Desktop\\output\\volume_data.xlsx"
]
output_folder_path = "C:\\Users\\f2021\\Desktop\\output"

# Counter to track which input file is being processed
file_counter = 0

def create_split_data_frame(root):
    global file_counter

    frame = customtkinter.CTkFrame(root, corner_radius=0, fg_color=get_current_fg_color())
    frame.grid_columnconfigure(0, weight=1)
    frame.grid_rowconfigure(8, weight=1)

    input_file_var = customtkinter.StringVar(value=input_file_paths[file_counter])
    output_folder_var = customtkinter.StringVar(value=output_folder_path)
    category_vars = {}

    def scan_and_create():
        nonlocal category_vars
        category_vars = scan_and_create_checkboxes(frame, input_file_var, categories_frame)

    customtkinter.CTkLabel(frame, text="Input Excel file:").grid(row=0, column=0, padx=20, pady=(20, 5), sticky="ew")
    input_file_label = customtkinter.CTkLabel(frame, text=f"Selected file: {os.path.basename(input_file_var.get())}", wraplength=400)
    input_file_label.grid(row=1, column=1, padx=20, pady=5, sticky="w")

    customtkinter.CTkLabel(frame, text="Output folder:").grid(row=2, column=0, padx=20, pady=(20, 5), sticky="ew")
    output_folder_label = customtkinter.CTkLabel(frame, text=f"Selected folder: {os.path.basename(output_folder_var.get())}", wraplength=400)
    output_folder_label.grid(row=3, column=1, padx=20, pady=5, sticky="w")

    categories_frame = customtkinter.CTkScrollableFrame(frame)
    categories_frame.grid(row=4, column=0, columnspan=2, padx=20, pady=20, sticky="nsew")

    scan_button = customtkinter.CTkButton(frame, text="Scan and Create Checkboxes", command=scan_and_create, width=150)
    scan_button.grid(row=5, column=0, columnspan=2, padx=20, pady=20, sticky="sew")

    sort_button = customtkinter.CTkButton(frame, text="Sort Selected Categories",
                                          command=lambda: sort_selected_categories(root, input_file_var.get(), output_folder_var.get(), category_vars), width=150)
    sort_button.grid(row=6, column=0, columnspan=2, padx=20, pady=20, sticky="sew")

    return frame

def scan_and_create_checkboxes(frame, input_file_var, categories_frame):
    input_file = input_file_var.get()
    if not input_file:
        messagebox.showerror("Error", "Please select an input file first")
        return

    try:
        data = pd.read_excel(input_file)
        categories = data['Category'].unique()

        for widget in categories_frame.winfo_children():
            widget.destroy()

        category_vars = {}
        row, col = 0, 0
        for category in categories:
            var = customtkinter.StringVar(value="0")
            checkbox = customtkinter.CTkCheckBox(categories_frame, text=category, variable=var)
            checkbox.grid(row=row, column=col, padx=10, pady=5, sticky="w")
            category_vars[category] = var

            col += 1
            if col == 2:
                col = 0
                row += 1

        return category_vars

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while scanning the file: {e}")

def sort_selected_categories(root, input_file, output_folder, category_vars):
    global file_counter

    if not input_file or not output_folder:
        messagebox.showerror("Error", "Please select both input file and output folder")
        return

    try:
        # Determine output file names based on the input file being processed
        if "area_data" in input_file:
            sorted_output_path = os.path.join(output_folder, "Sorted_Area_Data.xlsx")
            cleaned_output_path = os.path.join(output_folder, "Cleaned_Sorted_Area_Data.xlsx")
        else:
            sorted_output_path = os.path.join(output_folder, "Sorted_Volume_Data.xlsx")
            cleaned_output_path = os.path.join(output_folder, "Cleaned_Sorted_Volume_Data.xlsx")

        data = pd.read_excel(input_file)
        selected_categories = [category for category, var in category_vars.items() if var.get() == "1"]
        category_data = {category: data[data['Category'].str.contains(category, case=False, na=False)] for category in selected_categories}

        with pd.ExcelWriter(sorted_output_path, engine='xlsxwriter') as writer:
            for category, df in category_data.items():
                df.to_excel(writer, sheet_name=category, index=False)

        sorted_excel_data = pd.ExcelFile(sorted_output_path)
        cleaned_data = {}
        for sheet_name in sorted_excel_data.sheet_names:
            df = pd.read_excel(sorted_output_path, sheet_name=sheet_name)
            cleaned_df = df.dropna(axis=1, how='all')
            cleaned_data[sheet_name] = cleaned_df

        try:
            with pd.ExcelWriter(cleaned_output_path, engine='xlsxwriter') as writer:
                for sheet_name, df in cleaned_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            messagebox.showinfo("Success", f"Sorted data has been saved to {sorted_output_path}\nCleaned data has been saved to {cleaned_output_path}")
        except PermissionError:
            alt_cleaned_output_path = cleaned_output_path.replace("Cleaned_", "Alt_Cleaned_")
            with pd.ExcelWriter(alt_cleaned_output_path, engine='xlsxwriter') as writer:
                for sheet_name, df in cleaned_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            messagebox.showinfo("Success", f"Sorted data has been saved to {sorted_output_path}\nCleaned data has been saved to {alt_cleaned_output_path}")

        # Move to the next input file
        file_counter += 1
        if file_counter < len(input_file_paths):
            root.destroy()
            start_application()  # Start the application again for the next file
        else:
            messagebox.showinfo("Info", "All files have been processed.")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def start_application():
    root = customtkinter.CTk()
    root.title("Excel Category Sorter")
    split_data_frame = create_split_data_frame(root)
    split_data_frame.pack(fill="both", expand=True)
    root.mainloop()

# Start the application for the first file
start_application()
