"""
Program: TXT Extract To Excel
Author: David J M Ellwood
Email: david@ellwood.email
Date: March 6, 2023

Description: This program reads in text files from a specified directory, applies data cleaning and transformation, and
exports the cleaned data to Excel files in a specified output directory. The program also compares the number of lines in
the original file to the new Excel file and writes any differences to an errors.txt file in the output directory.

Change Control:
Version 1.0 - Initial release
"""

# Import the necessary libraries
import pandas as pd   # for working with data in tabular format
from pathlib import Path   # for working with file paths

# Define the input and output directories
input_dir = Path(r'C:\Users\David\OneDrive\Documents\SAP\Automations\SAP Extract To Excel\SAP Extract Input')
output_dir = input_dir.parent / 'Excel Output'

# Loop through each file in the input directory
for file_path in input_dir.glob('*.txt'):
    
    # Read the input file as a pandas DataFrame
    df = pd.read_csv(file_path,
                     engine='python',   # specify the engine to use for reading the file
                     sep='|',   # specify the delimiter used in the file
                     quotechar='"',   # specify the quote character used in the file
                     encoding='ISO-8859-1',   # specify the character encoding of the file
                     infer_datetime_format=True,   # infer the datetime format from the data
                     parse_dates=True,   # parse the data as datetime objects
                     dayfirst=False,   # specify the format of the date string
                     on_bad_lines='skip')   # skip lines that cannot be parsed as data
    
    # Replace all instances of "01.01.0001" with "01/01/1900" in the DataFrame
    df = df.replace('01.01.0001', '01/01/1900')
    
    # Define a custom date parser function to parse date strings in the format DD.MM.YYYY
    def date_parser(val):
        if isinstance(val, str):
            if len(val) == 10:   # check if the date string is in the expected format
                val = val.replace('.', '/')   # replace periods with slashes in the date string
                try:
                    return pd.to_datetime(val, format='%d/%m/%Y').date()   # parse the date string as a datetime object
                except ValueError:
                    pass
        return val   # return the original value if it cannot be parsed as a date string
    
    # Apply the custom date parser function to all values in the DataFrame
    df = df.applymap(date_parser)
    
    # Save the DataFrame as an Excel file in the output directory
    excel_file_path = output_dir / (file_path.stem + '.xlsx')
    df.to_excel(excel_file_path, index=False)
    
    # Check if the number of lines in the original file matches the number of lines in the new file
    original_lines = sum(1 for line in open(file_path, encoding='ISO-8859-1'))
    new_lines = len(df.index)
    
    # If the number of lines does not match, write the details to an errors.txt file in the output directory
    if original_lines != new_lines:
        
        # Check for missing lines between the original and new files
        with open(file_path, encoding='ISO-8859-1') as f:
            original_data = f.readlines()
        new_data = df.to_csv(index=False, header=None, sep='|').split('\n')
        missing_lines = [i for i, line in enumerate(original_data) if line.strip() != new_data[i].strip()]

        # Write the error details to the errors.txt file
        error_file = output_dir / 'errors.txt'
        with open(error_file, 'a', encoding='ISO-8859-1') as f:
            f.write(f'{file_path.name}: {original_lines} lines in original file, {new_lines} lines in new file\n')
            if missing_lines:
                f.write('Missing lines:\n')
                for line in missing
