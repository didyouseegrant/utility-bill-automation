import os
import re

# Define the directory containing the PDF files
directory = r"//path//to//your//input//folder"

# List of full account numbers - these are placeholders
account_numbers = [
    "1111111111-1", "2222222222-2", "3333333333-3", "4444444444-4", "5555555555-5",
    "6666666666-6", "7777777777-7", "8888888888-8", "9999999999-9", "1010101010-0",
    "1212121212-1", "1313131313-2", "1414141414-3", "1515151515-4", "1616161616-5",
    "1717171717-6", "1818181818-7", "1919191919-8", "2020202020-9", "2121212121-0",
    "2222222222-1", "2323232323-2", "2424242424-3", "2525252525-4", "2626262626-5",
    "2727272727-6", "2828282828-7", "2929292929-8", "3030303030-9"
]

# By default, PG&E bills are downloaded only including last four digits of account number
# Create a dictionary mapping the last four digits to full account numbers
last_four_to_full = {acc[-6:-2]: acc for acc in account_numbers}

# Iterate through files in the directory
for filename in os.listdir(directory):
    old_path = os.path.join(directory, filename)

    # Ensure it's a PDF file
    if os.path.isfile(old_path) and filename.lower().endswith(".pdf"):
        # Extract the last four digits from the filename using regex
        match = re.search(r"(\d{4})custbill", filename)
        if match:
            last_four_digits = match.group(1)

            # Find the matching account number
            if last_four_digits in last_four_to_full:
                new_filename = f"{last_four_to_full[last_four_digits]}.pdf"
                new_path = os.path.join(directory, new_filename)

                # Rename the file if it has a valid match
                if new_filename != filename:
                    os.rename(old_path, new_path)
                    print(f"Renamed: {filename} -> {new_filename}")

print("Renaming process completed.")
