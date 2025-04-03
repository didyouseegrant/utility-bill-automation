import fitz  # PyMuPDF for PDF text extraction
import re  # Regular expressions for pattern matching
from openpyxl import load_workbook  # For handling Excel files
import os

# Define paths
pdf_folder = r"F:\Accounts Payable\Utilities\Marin Sanitary Service\2025\Bills\02. February"
excel_file = r"F:\Accounts Payable\Utilities\Marin Sanitary Service\2025\MSS 0225.xlsx"
sheet_name = "Monthly Sorted by Prop"  # Updated sheet name

# List of account numbers in the required order
pdf_order = [
    "51925", "51931", "51932", "51933", "51934", "51935", "51937", "51939",
    "51940", "51941", "51942", "51943", "51944", "51945", "51946", "47569",
    "47570", "47572", "47566", "47567", "47571", "47563", "47565", "47564",
    "48057", "48058", "48056", "48055", "48054", "48053", "48051", "102324",
    "102323", "2743", "48059", "48060", "48061", "48062", "2502", "2503",
    "2504", "2486", "2488", "2493", "2489", "2481", "2482", "2484", "2485",
    "2492", "2483", "2491", "2495", "2496", "2499", "2501", "2505", "2497",
    "2695", "2521", "40852", "42847", "2500", "2490", "114265", "101779",
    "27922", "2494"
]

# Get all PDF files in the folder
pdf_files = os.listdir(pdf_folder)

# Filter and sort PDFs based on account numbers
sorted_pdfs = []
for account_number in pdf_order:
    matched_files = [file for file in pdf_files if account_number in file and file.endswith(".pdf")]
    sorted_pdfs.extend(matched_files)  # Maintain order from pdf_order

# Compile a regex pattern to extract "Total Due" amounts (handles numbers with or without commas)
amount_pattern = re.compile(r"TOTAL\s+DUE\s+(\d{1,3}(?:,\d{3})*\.\d{2}|\d+\.\d{2})", re.IGNORECASE)

# Open the Excel workbook
wb = load_workbook(excel_file)
ws = wb[sheet_name]

# Start writing data to Excel from row 4 onwards
current_row = 4

# Iterate through the sorted PDFs
for pdf_file in sorted_pdfs:
    pdf_path = os.path.join(pdf_folder, pdf_file)

    doc = fitz.open(pdf_path)
    text = ""

    # Extract text from each page
    for page in doc:
        text += page.get_text()
    doc.close()

    # Debugging: Print the extracted text
    # print(f"Extracted text from {pdf_file}:\n{text}\n{'-'*50}")

    # Search for the 'Total Due' amount
    match = amount_pattern.search(text)
    if match:
        amount_due = match.group(1)  # Extract the amount
        print(f':D Found amount in {pdf_file}: {amount_due}')
        
        # Special case for accounts 27922 and 2494
        if "27922" in pdf_file:
            ws.cell(row=72, column=6, value=amount_due)  # Row 72, Column F (Excel column 6)
        elif "2494" in pdf_file:
            ws.cell(row=75, column=6, value=amount_due)  # Row 75, Column F (Excel column 6)
        else:
            ws.cell(row=current_row, column=6, value=amount_due)  # Insert into next available row

    else:
        print(f'❌ No amount found in {pdf_file}')

    # Move to the next row, except for special cases
    if "27922" not in pdf_file and "2494" not in pdf_file:
        current_row += 1
  

# Save the updated Excel file
wb.save("F:\\Accounts Payable\\Utilities\\Marin Sanitary Service\\2025\\test.xlsx")
print(":# Data extraction and insertion completed successfully.")
