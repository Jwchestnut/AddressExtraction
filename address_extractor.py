# Version: 1.0
# Description: Initial version of the address extractor script. This version includes a manual approach to identify street addresses by finding street indicators and capturing relevant information.
# Manual approach to identify the street address by finding the street indicator and capturing everything to the left through the first set of numerical digits followed by a space

# List of street indicators
street_indicators = ["Road", "Rd.", "Street", "St.", "Court", "Circle", "Ave.", "Avenue", "Blvd.", "Boulevard", "Drive", "Dr."]

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

# Example usage
first_line = "#573 Jeanette John owner 10 Mountain Laurels Drive #101 Nashua NH 03062"
manual_street_address = extract_address(first_line)
print("Extracted Address:", manual_street_address)
