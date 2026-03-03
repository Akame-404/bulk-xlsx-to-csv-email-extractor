# Bulk XLSX to CSV Email Extractor

## Overview

Bulk XLSX to CSV Email Extractor is a Windows-based automation tool designed to:

- Convert all `.xlsx` files in a directory to `.csv`
- Extract email addresses from generated CSV files
- Remove duplicates
- Generate a consolidated `emails.csv` output file

This tool is useful for data processing, automation workflows, and structured email extraction tasks.

---

## Features

- Bulk Excel to CSV conversion (COM automation)
- Regex-based email extraction
- Automatic deduplication
- UTF-8 output encoding
- Zero external dependencies (native Windows + Excel)

---

## Requirements

- Windows 10 / 11
- Microsoft Excel installed
- PowerShell enabled

---

## How It Works

1. Iterates through all `.xlsx` files in the script directory.
2. Converts them to `.csv` using Excel COM objects.
3. Scans all CSV files (excluding `emails.csv`).
4. Extracts email patterns using regex: [a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}
5. Sorts and removes duplicates.
6. Outputs a clean `emails.csv` file.

---

## Usage

1. Place the `.bat` script in a folder containing `.xlsx` files.
2. Run the script.
3. The following will be generated:
- Converted `.csv` files
- A consolidated `emails.csv` file

---

## Output Format

`emails.csv`
Email
example1@domain.com
example2@domain.com
...

---

## Security & Compliance Notice

This tool extracts email addresses from local files only.

Users are responsible for ensuring compliance with:
- GDPR
- Local data protection regulations
- Organizational data policies

---

## Potential Improvements

- Cross-platform Python version
- Logging system
- Input validation
- CLI parameters
- Multithreaded processing
