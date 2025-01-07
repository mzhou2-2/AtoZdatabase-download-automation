import os
import pandas as pd
from pathlib import Path

def combine_excel_files(folder_path, output_file):
    """
    Combines all Excel files in the specified folder into a single Excel file.

    Parameters:
    - folder_path (str): The path to the folder containing Excel files.
    - output_file (str): The path for the combined output Excel file.
    """
    # Create a Path object
    folder = Path(folder_path)

    # Check if the folder exists
    if not folder.exists() or not folder.is_dir():
        raise ValueError(f"The folder path {folder_path} does not exist or is not a directory.")

    # Supported Excel file extensions
    excel_extensions = ['.xlsx', '.xls', '.xlsm']

    # List to hold individual DataFrames
    df_list = []

    # Iterate over all Excel files in the folder
    for file in folder.iterdir():
        if file.is_file() and file.suffix.lower() in excel_extensions:
            try:
                print(f"Reading file: {file.name}")
                
                # Read the Excel file
                # If there are multiple sheets and you want to read a specific one, specify sheet_name
                df = pd.read_excel(file, sheet_name=0)  # Reads the first sheet by default
                
                # Optionally, add a column to identify the source file
                df['Source_File'] = file.name
                
                df_list.append(df)
            except Exception as e:
                print(f"Error reading {file.name}: {e}")

    if not df_list:
        print("No Excel files found to combine.")
        return

    # Concatenate all DataFrames
    combined_df = pd.concat(df_list, ignore_index=True)

    # Optional: Remove duplicate rows
    combined_df.drop_duplicates(inplace=True)

    # Optional: Reset index
    combined_df.reset_index(drop=True, inplace=True)

    # Save the combined DataFrame to a new Excel file
    try:
        combined_df.to_excel(output_file, index=False)
        print(f"Combined data saved to {output_file}")
    except Exception as e:
        print(f"Error saving the combined Excel file: {e}")

if __name__ == "__main__":
    # Specify the folder containing Excel files
    folder_path = r"C:\Users\matth\OneDrive\Desktop\AtoZ_SC_all"

    # Specify the output file path
    output_file = r"C:\Users\matth\OneDrive\Desktop\AtoZ_SC_all\Combined_AtoZ_SC_All.xlsx"

    combine_excel_files(folder_path, output_file)
