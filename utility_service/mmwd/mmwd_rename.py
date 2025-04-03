import os
import re

# Define the folder containing the PDFs
folder_path = r"\\main-2019\rf$\gchrisman\Documents\test2"

# Define the list of new names
new_names = [
    573517, 573520, 573521, 573523, 573525, 573526, 573527, 573528, 573530, 
    573531, 573532, 573534, 573535, 573536, 573537, 573538, 573539, 573540, 
    573541, 573542, 573543, 573544, 573547, 573548, 573585, 573586, 573587, 
    573589, 573590, 573592, 573594, 573595, 573597, 573598, 573599, 573600, 
    573601, 573602, 573603, 573604, 573605, 573606, 573607, 573608, 573609, 
    573610, 573611, 573612, 573613, 573614, 573615, 573617, 573618, 573619, 
    573620, 573621, 573622, 573623, 573624
]

# Rename "Doc.pdf" to "Doc (0).pdf" if it exists
doc_path = os.path.join(folder_path, "Doc.pdf")
doc_renamed_path = os.path.join(folder_path, "Doc (0).pdf")

if os.path.exists(doc_path):
    os.rename(doc_path, doc_renamed_path)
    print(f'Renamed "Doc.pdf" → "Doc (0).pdf"')

# Function to extract the number from a filename (handles "Doc", "Doc (1)", etc.)
def extract_number(filename):
    match = re.search(r'\((\d+)\)', filename)  # Find number inside parentheses
    return int(match.group(1)) if match else 0  # Return 0 if no number found (for "Doc (0)")

# Get a sorted list of all PDF files in the folder using natural sorting
pdf_files = sorted(
    [f for f in os.listdir(folder_path) if f.startswith("Doc") and f.endswith(".pdf")],
    key=extract_number  # Sort numerically by extracted number
)

# Ensure we have the same number of PDFs and names
if len(pdf_files) != len(new_names):
    print("Error: The number of PDFs and the number of new names do not match.")
else:
    for old_name, new_name in zip(pdf_files, new_names):
        old_path = os.path.join(folder_path, old_name)
        new_path = os.path.join(folder_path, f"{new_name}.pdf")

        os.rename(old_path, new_path)
        print(f"Renamed: {old_name} → {new_name}.pdf")

print("Renaming complete.")
