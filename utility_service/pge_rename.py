import os
import re

# Define the directory containing the PDF files
directory = r"\\main-2019\rf$\gchrisman\Documents\test2"

# List of full account numbers
account_numbers = [
    "1051232522-5", "1947775177-0", "7656100844-2", "1425713638-3", "8176683824-7", 
    "3937199847-5", "8478299735-1", "9655656259-1", "7447349142-6", "9010017104-1",
    "0270611640-3", "8947324014-4", "7947324078-0", "1739465846-6", "6447538128-7",
    "1906082438-0", "4783150844-6", "3228944784-2", "2697323371-6", "2531108473-2",
    "1467300186-4", "2822775121-2", "1417723039-9", "0937554826-1", "0406127383-7",
    "2176684208-8", "9739460119-5", "2020219221-5", "1852724251-7"
]

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
