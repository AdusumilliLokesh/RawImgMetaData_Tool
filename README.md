# Raw Image Metadata Generator for GNP (Georgia Newspaper Project)

This repository contains an open-source Python script developed for automating the generation of raw image metadata for the **Georgia Newspaper Project (GNP)**, a division of the **Digital Library of Georgia (DLG)**. It is designed to streamline the production of structured Excel metadata from raw issue-level data.

This tool is free and open-source for public use, with the understanding that **official support is exclusively available to members and affiliates of the Digital Library of Georgia (DLG)**.

---

## Features

* Converts raw Excel data of newspaper issues into a structured metadata sheet
* Auto-generates metadata rows for each page based on configuration
* Flexible configuration via `config.json`
* Outputs fully formatted `.xlsx` metadata files with standardized fields

---

## Requirements

* Python 3.7+
* [pandas](https://pandas.pydata.org/)
* [openpyxl](https://openpyxl.readthedocs.io/)

You can install dependencies using:

```bash
pip install pandas openpyxl
```

---

## Installation

Download the PDF: Installation Instructions for setting up the environment and installing required libraries using PyCharm.

---

## Usage

Download the PDF: Usage Instructions to:

* Configure your `config.json`
* Prepare your raw Excel input file
* Run the script and generate metadata output

---

## Auditing the Output

After generating your metadata file, please consult the **Audit Manual PDF** included in the `docs/` folder to ensure the accuracy of your output. This guide walks you through:

- Verifying page count against scanned files  
- Validating dates, volume numbers, and issue numbers  
- Correcting input errors and regenerating the final metadata file  

**Note:** Auditing is highly recommended before finalizing or submitting any metadata files for archiving or publication.
---


## Repository Structure

```
project_root/
├── main.py                # Main executable script
├── config.json            # Configuration file for metadata and paths
├── IO Folder/             # Contains input and output Excel files
│   ├── Raw_OEW_2023.xlsx  # Sample input file
│   └── OEW_2023_metadata.xlsx  # Output file
├── docs/
│   ├── Installation_Instructions.pdf
│   └── Usage_Instructions.pdf
|   └──  Audit Manual for Metadata Output Validation.pdf
└── README.md              # This file
```

---

## License

This project is released as open-source for educational and archival purposes. Usage by external parties is permitted, but **support and maintenance are only guaranteed for internal DLG projects.**

For questions related to usage within DLG workflows, contact the internal DLG metadata team.
