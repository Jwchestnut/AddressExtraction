# AddressExtraction
Identify a USA address within other text in an Excel field. Copy to specific address columns.

## address_extractor.py (Version 1.0)
### Purpose
This script is designed to extract street addresses from a given line of text. It is a crucial part of the address extraction project, specifically targeting USA addresses within Excel fields.

### Functionality
By utilizing a list of common street indicators, the script identifies and extracts the full address. The main function, `extract_address`, takes a line of text and returns the extracted address or a "No match found" message if no address is detected.

### Unique Features
The script employs a manual approach, allowing for flexibility in handling various address formats. It can be easily extended or modified to accommodate specific needs.

### Considerations
As this is the initial version (1.0), future updates may include optimizations, support for additional address formats, and integration with other components of the project.

## Updated_Lot_Cards_Script.py (Version 1.0)
### Purpose
This script is designed to update an Excel file with extracted addresses, ZIP codes, cities, and states. It is a crucial part of the address extraction project, specifically targeting lot cards.

### Functionality
By utilizing the `extract_address` function from `address_extractor.py`, the script extracts street addresses from the 'Original_Address' column. It also extracts ZIP codes using regex and queries the corresponding city and state using a ZIP code mapping file.

### Unique Features
The script integrates with `address_extractor.py` and handles Excel workbook operations, allowing for automated updates to lot cards with accurate address information.

### Considerations
As this is the initial version (1.0), future updates may include optimizations, support for additional Excel formats, and integration with other components of the project.
