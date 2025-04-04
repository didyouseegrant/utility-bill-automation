# Utility Bill Automation System

This project automates the processing of monthly utility bills for three major service providers: Marin Municipal Water District (MMWD), Marin Sanitary Service (MSS), and PG&E. It streamlines downloading, renaming, parsing, and organizing over 160 utility bills each month using Python.

---

## Features

- **Automated PDF downloads** via Selenium (MMWD & PG&E)
- **File renaming** based on account numbers for easier tracking
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
│   │   ├── mmwd_download.py
│   │   ├── mmwd_rename.py
│   │   ├── mmwd_bills.py
│   │   └── mmwd_analysis.py
│   ├── mss/
│   │   ├── mss_bills.py
│   │   └── mss_analysis.py
│   ├── pge/
│   │   ├── pge_download.py
│   │   ├── pge_rename.py
│   │   └── pge_analysis.py
├── sample_data/
│   ├── sample_bills/
│   │   ├── mmwd_sample_bill.pdf
│   │   └── mss_sample_bill.pdf
│   └── sample_output.xlsx
└── README.md
```

