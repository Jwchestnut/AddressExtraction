from openpyxl import load_workbook
import re
import pandas as pd
import sys

# Add the path to the directory where address_extractor.py is located
path_to_address_extractor = r"C:\MAC-009 Test\MAC-009 Files"
sys.path.append(path_to_address_extractor)

from address_extractor import extract_address

# Define the file path for the target Excel file
file_path = r"C:\MAC-009 Test\MAC-009 Files\Test4.xlsx"

# Path to the ZIP code mapping file
zip_mapping_file = r"C:\MAC-009 Test\MAC-009 Files\ZIP_Locale_Detail.xls"

# Read the ZIP code mapping file into a DataFrame
zip_mapping_df = pd.read_excel(zip_mapping_file)

# Convert the 'PHYSICAL ZIP', 'PHYSICAL CITY', and 'PHYSICAL STATE' columns into a dictionary
zip_mapping = dict(zip(zip_mapping_df['PHYSICAL ZIP'], zip_mapping_df['PHYSICAL CITY'] + ", " + zip_mapping_df['PHYSICAL STATE']))

# Load the Excel workbook
workbook = load_workbook(filename=file_path)

# Access the specific worksheet
worksheet = workbook.active

# Iterate through column C (assuming the data starts from row 1, skipping the header row)
for row, cell in enumerate(worksheet['C'][1:]):
    # Extract the street address using the imported function
    address = extract_address(str(cell.value))
    
    # Write the street address to the corresponding cell in column D
    worksheet[f'D{row + 2}'] = address

    # Extract the ZIP code using regex
    zip_code_match = re.search(r' \d{5}', str(cell.value))
    if zip_code_match and not str(cell.value).startswith("#"):
        zip_code = zip_code_match.group(0).strip()
        
        # Write ZIP code to column E
        worksheet[f'E{row + 2}'] = zip_code
        
        # Query city and state using ZIP code and write to column F
        city_state = zip_mapping.get(int(zip_code), 'Not found')
        worksheet[f'F{row + 2}'] = city_state

# Save the modified Excel file
output_path = r"C:\MAC-009 Test\MAC-009 Files\Modified_Test4.xlsx"
workbook.save(filename=output_path)
