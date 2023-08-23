from openpyxl import load_workbook
import re
import pandas as pd
import sys

# Add the path to the directory where address_extractor.py is located
path_to_address_extractor = r"C:\MAC-009 Test\MAC-009 Files"
sys.path.append(path_to_address_extractor)

from address_extractor import extract_address

# Define the file path for the target Excel file
file_path = r"C:\MAC-009 Test\MAC-009 Files\Lot Cards - Trilagen.xlsx"

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

# Iterate through the 'Original_Address' column (column D, skipping the header row)
for row, cell in enumerate(worksheet['D'][1:]):
    # Extract the street address using the imported function
    address = extract_address(str(cell.value))
    
    # Write the street address to the corresponding 'Street' cell in column M
    worksheet[f'M{row + 2}'] = address

    # Extract the ZIP code using regex (ignoring strings that start with "#")
    zip_code_match = re.search(r' \d{5}\b', str(cell.value))
    if zip_code_match and not str(cell.value).startswith("#"):
        zip_code = zip_code_match.group(0).strip()
        
        # Write ZIP code to 'Zip' cell in column N
        worksheet[f'N{row + 2}'] = zip_code
        
        # Query city and state using ZIP code and write to 'City' and 'State' cells in columns O and P
        city_state = zip_mapping.get(int(zip_code), 'Not found').split(", ")
        worksheet[f'O{row + 2}'] = city_state[0] # City
        worksheet[f'P{row + 2}'] = city_state[1] # State

# Save the modified Excel file
output_path = r"C:\MAC-009 Test\MAC-009 Files\Modified_Lot_Cards_Trilagen.xlsx"
workbook.save(filename=output_path)
