import subprocess
import sys
import time
from tkinter import Tk
from tkinter.filedialog import askopenfilenames
from tkinter import messagebox

def run_script(script_path):
    try:
        result = subprocess.run([sys.executable, script_path], check=True, capture_output=True, text=True)
        print(f"Successfully ran {script_path}")
        print(result.stdout)
    except subprocess.CalledProcessError as e:
        print(f"Error running {script_path}")
        print(e.stderr)
        sys.exit(1)

def run_scripts_in_sequence(scripts):
    Tk().withdraw()  # Prevents the root window from appearing
    messagebox.showinfo("Information", "Running all provided scripts in sequence.")

    # Run all scripts in sequence
    for script in scripts:
        run_script(script)
        time.sleep(1)  # Delay of 1 second after each script

    print("All scripts executed successfully.")
    messagebox.showinfo("Success", "All scripts executed successfully.")

# Example usage, to be removed when importing this function in other scripts
if __name__ == "__main__":
    # List of Python scripts to run (hardcoded paths)
    scripts = [
        r"C:\\Users\\f2021\\Desktop\\Rajiv_Sir_Project\\CostPredictor\\CostPredictor\\backend_for_single_task\\sorting_excel_data.py",
        r"C:\\Users\\f2021\\Desktop\\Rajiv_Sir_Project\\CostPredictor\\CostPredictor\\backend_for_single_task\\demerging_area_volume.py",
        r"C:\\Users\\f2021\\Desktop\\Rajiv_Sir_Project\\CostPredictor\\CostPredictor\\backend_for_single_task\\split_data_frame.py",
        r"C:\\Users\\f2021\\Desktop\\Rajiv_Sir_Project\\CostPredictor\\CostPredictor\\backend_for_single_task\\reformat.py",
        r"C:\\Users\\f2021\\Desktop\\Rajiv_Sir_Project\\CostPredictor\\CostPredictor\\backend_for_single_task\\quantity_calculator.py",
        r"C:\\Users\\f2021\\Desktop\\Rajiv_Sir_Project\\CostPredictor\\CostPredictor\\backend_for_single_task\\display_chart.py",
        r"C:\\Users\\f2021\\Desktop\\Rajiv_Sir_Project\\CostPredictor\\CostPredictor\\backend_for_single_task\\cleaning_file.py"
    ]

    run_scripts_in_sequence(scripts)
