# Utility Bill Automation System

This project automates the processing of monthly utility bills for three major service providers: Marin Municipal Water District (MMWD), Marin Sanitary Service (MSS), and PG&E. It streamlines downloading, renaming, parsing, and organizing over 160 utility bills each month using Python and Excel.

For a more detailed look into the project, check out my [Notion case study!](https://www.notion.so/Utility-Bill-Automation-System-Python-Excel-1c9fcf2408d080d3a433f4c094d172f7)

---

## Features

- **Automated PDF downloads** via Selenium (MMWD & PG&E)
- **File renaming** based on account numbers for easier tracking (MMWD & PGE)
- **PDF parsing** using PyMuPDF to extract "Total Amount Due" (MMWD & MSS) & "Total Gallon Used" (MMWD)
- **Excel integration** with openpyxl to log and organize bill data
- **Property-level categorization** across 25 managed properties

---

## Tools & Libraries

- Python 3  
- Selenium  
- PyMuPDF (`fitz`)  
- openpyxl  
- ChromeDriver  
- Excel

---

## Structure

```plaintext
├── utility_service/
│   ├── mmwd/
│   │   ├── 1. mmwd_download.py
│   │   ├── 2. mmwd_rename.py
│   │   ├── 3. mmwd_bills.py
│   │   └── 4. mmwd_analysis.py
│   ├── mss/
│   │   ├── 1. mss_bills.py
│   │   └── 2. mss_analysis.py
│   ├── pge/
│   │   ├── 1. pge_download.py
│   │   ├── 2. pge_rename.py
│   │   └── 3. pge_analysis.py
├── sample_data/
│   ├── sample_bills/
│   │   ├── mmwd_sample_bill.pdf
│   │   └── mss_sample_bill.pdf
│   ├── mmwd_sample_bill.xlsx
|   ├── mss_sample_bill.xlsx
|   └── pge_sample_sheet.xlsx
└── README.md
```

