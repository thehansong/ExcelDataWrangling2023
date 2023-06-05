""" *****************************************************************************
-*- coding: utf-8 -*-
    \file           3.Temperature_Analysis.py
    \author         Hans (You Yang) ONG
    
    \creation date  050623
    \last updated   050623

    \brief          This script performs temperature analysis on merged files.
                    It checks the temperature differences within each file and
                    highlights any rows that exceed the specified threshold.
***************************************************************************** """
import pandas as pd
from colorama import Fore, Style
import os
import re

# Path to the merged files directory
merged_files_directory = 'Merged/'

# Get a list of all files in the merged files directory
files = os.listdir(merged_files_directory)

# Sort the files based on the numeric part of their names
files = sorted(files, key=lambda x: int(re.search(r'\d+', x).group()))

# Define the temperature columns to compare
temperature_columns = ['Temp', 'Temp_Blaze_Stats', 'Temp_Blaze_LW_Dist', 'Temp_Blaze_CW_Dist']

# Set the temperature difference threshold (in percentage)
threshold = 0.1 #%

# Iterate over the files and check temperature differences
for filename in files:
    print(f"Checking file: {filename}")
    print("------------------------------------------")

    # Read the merged file
    merged_file = pd.read_excel(merged_files_directory + filename)

    # Create a new DataFrame to store the problem rows and the additional columns
    problem_rows = pd.DataFrame(columns=['Local Time'] + temperature_columns + ['Percentage Difference'])
    
    # Iterate over the rows and calculate the percentage difference
    for index, row in merged_file.iterrows():
        temperatures = row[temperature_columns]
        max_diff = max(temperatures) - min(temperatures)
        percentage_diff = (max_diff / min(temperatures)) * 100
        
        # Check if the percentage difference exceeds the threshold
        if percentage_diff > threshold:
            problem_row = [row['Local Time']] + row[temperature_columns].tolist() + [percentage_diff]
            problem_rows.loc[index] = problem_row
            print(f"{Fore.RED}Local Time: {row['Local Time']} - Row {index+1}: Temperature difference greater than {threshold}%{Style.RESET_ALL}")

    if not problem_rows.empty:
        # Save the problem rows to a new Excel file
        new_filename = f"ProblemRows_{filename}"
        problem_rows.to_excel(new_filename, index=False)
        print(f"Problem rows are saved in {new_filename}")
    else:
        print(f"All temperature differences are within the {threshold}% range.")

    print("------------------------------------------\n")
