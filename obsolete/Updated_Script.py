
from openpyxl import load_workbook
import re
import sys

# Add the path to the directory where address_extractor.py is located
path_to_address_extractor = r"C:\MAC-009 Test\MAC-009 Files"
sys.path.append(path_to_address_extractor)

from address_extractor import extract_address  # Import the function from the address_extractor.py file

# Define the file path for the target Excel file
file_path = r"C:\MAC-009 Test\MAC-009 Files\Test4.xlsx"

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

    # Extract the ZIP code using regex, ignoring strings that start with #
    zip_code_match = re.search(r'(?<![#])\b\d{5}\b', str(cell.value))
    
    if zip_code_match:
        zip_code = zip_code_match.group(0)
        # Write ZIP code to column E
        worksheet[f'E{row + 2}'] = zip_code

# Save the modified Excel file
output_path = r"C:\MAC-009 Test\MAC-009 Files\Modified_Test4.xlsx"
workbook.save(filename=output_path)
