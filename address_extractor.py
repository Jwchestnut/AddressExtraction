# Version: 1.1
# Description: Initial version of the script for updating lot cards with extracted addresses, ZIP codes, cities, and states.
feature/zip_code
# List of street indicators
street_indicators = ["Road", "Rd.", "Street", "St.", "Court", "Ct.", "Circle", "Cir.", "Avenue", "Ave.", "Boulevard", "Blvd.", "Drive", "Dr.", "Lane", "Ln.", "Place", "Pl.", "Terrace", "Ter.", "Highway", "Hwy.", "Parkway", "Pkwy.", "Alley", "Aly.", "Freeway", "Fwy.", "Expressway", "Expy.", "Trail", "Trl.", "Way", "Way", "Square", "Sq.", "Loop", "Loop"]

# Function to extract the address using the manual approach
def extract_address(line):
    # Loop through the street indicators to find a match
    for indicator in street_indicators:
        # Find the position of the street indicator
        indicator_pos = line.find(indicator)
        if indicator_pos != -1:
            # Extract everything to the left of the street indicator
            left_part = line[:indicator_pos]
            # Split the left part into words and reverse the order
            words = left_part.split()[::-1]
            # Initialize the address with the street indicator
            address = [indicator]
            # Loop through the words, adding them to the address until a word containing only digits is found
            for word in words:
                if word.isdigit():
                    # Add the word containing only digits and stop
                    address.append(word)
                    break
                # Add other words to the address
                address.append(word)
            # Reverse the address to the original order and join it into a string
            address_str = " ".join(address[::-1])
            return address_str
    return "No match found"

# Rest of the code for updating lot cards with extracted addresses, ZIP codes, cities, and states
from openpyxl import load_workbook
import re
import pandas as pd


# Define the file path for the target Excel file
file_path = r"C:\MAC-009 Test\MAC-009 Files\AddressExtraction\Working Files\lotcards-trilagen.csv"

# Path to the ZIP code mapping file
zip_mapping_file = r"C:\MAC-009 Test\MAC-009 Files\AddressExtraction\Working Files\ZIP_Locale_Detail.xls"

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
feature/zip_code
    # Extract the street address using the extract_address function

    address = extract_address(str(cell.value))
    
    # Write the street address to the corresponding 'Street' cell in column M
    worksheet[f'M{row + 2}'] = address

feature/zip_code
    # Continue with the rest of the code as previously described...

# Save the modified Excel file
output_path = r"C:\MAC-009 Test\MAC-009 Files\AddressExtraction\Working Files\Modified_Lot_Cards_Trilagen.xlsx"
workbook.save(filename=output_path)


