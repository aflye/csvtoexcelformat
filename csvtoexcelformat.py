# -*- coding: utf-8 -*-
"""
Created on Tue Jan 23 16:56:03 2024

@author: ryan.wandera
"""
import os
import pandas as pd

# Define the maximum number of rows that Excel can handle
MAX_ROWS = 1048576

# Get a list of all files in the current directory
all_files = os.listdir('.')

# Iterate over the files
for filename in all_files:
    # Check if the file is a CSV file
    if filename.endswith('.csv'):
        # Read the CSV file using pandas
        df = pd.read_csv(filename)
        
        # Check if the number of rows in the DataFrame exceeds the maximum
        if len(df) > MAX_ROWS:
            # Split the DataFrame into two parts
            half = len(df) // 2
            df1, df2 = df[:half], df[half:]
            
            # Create new filenames by appending '1' and '2' to the original filename
            new_filename1 = filename[:-4] + '_1.xlsx'
            new_filename2 = filename[:-4] + '_2.xlsx'
            
            # Write the two DataFrames to separate Excel files
            df1.to_excel(new_filename1, index=False)
            df2.to_excel(new_filename2, index=False)
            
            # Print a success message
            print(f"Successfully split and converted {filename} to {new_filename1} and {new_filename2}")
        else:
            # Create a new name for the Excel file by replacing '.csv' with '.xlsx'
            new_filename = filename[:-4] + '.xlsx'
            
            # Write the DataFrame to an Excel file
            df.to_excel(new_filename, index=False)
            
            # Print a success message
            print(f"Successfully converted {filename} to {new_filename}")
