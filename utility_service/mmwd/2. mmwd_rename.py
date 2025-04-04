import os
import re

# Define the folder containing the PDFs
folder_path = r"path/to/your/input/folder"

# Define the list of new names
new_names = [
    111111, 222222, 333333, 444444, 555555, 666666, 777777, 888888, 999999, 
    101111, 112222, 123333, 134444, 145555, 156666, 167777, 178888, 189999, 
    191111, 202222, 213333, 224444, 235555, 246666, 257777, 268888, 279999, 
    281111, 292222, 303333, 314444, 325555, 336666, 347777, 358888, 369999, 
    371111, 382222, 393333, 404444, 415555, 426666, 437777, 448888, 459999, 
    461111, 472222, 483333, 494444, 505555, 516666
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
