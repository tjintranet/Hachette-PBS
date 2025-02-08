# POD Order Processing Tool

A web-based application that processes Excel files containing order information and converts them into a standardized CSV format for POD processing. Uploading the PBS file to the Hachette SFTP will force the download of the invoice and packing documents.

## Features

- Excel file upload and processing
- Interactive data preview table
- Row deletion (single or multiple)
- Standardized CSV output format
- Automatic courier service name standardization
- Line number generation
- Tracking reference generation

## Input Format

The tool expects an Excel file (.xlsx or .xls) with the following columns:
- Reference
- ISBN
- Quantity
- Courier
The Excel is downloaded from the following report in PACE - Search for the order/reference number
http://192.168.10.251/epace/company:c001/inquiry/UserDefinedInquiry/view/5284?

## Output Format

Generates a CSV file named `T1.M{reference}.PBS` with the following format:
```
reference,line_number,isbn,date,courier,quantity,status,tracking_reference
```

### Field Formatting

- **Line Numbers**: Zero-padded 5-digit numbers (e.g., "00001")
- **Date**: Fixed format DDMMYYYY
- **Courier**: Standardized names
  - "DPD - DPD" → "DPD"
  - "DHL - DHL Express" → "DHL Exp"
  - "Royal Mail - ROYAL M" → "ROYAL M"
- **Status**: Always set to "1"
- **Tracking Reference**: Format %0SL30HE1550 + random numbers
- **Quantity**: Defaults to "1" if not specified

## Usage

1. Open the application in a web browser
2. Click "Choose File" to select your Excel file
3. Review the processed data in the preview table
4. Use the checkboxes and delete buttons to remove any unwanted rows
5. Click "Download CSV" to generate and download the formatted CSV file

## Setup

The application requires the following dependencies (included via CDN):
- Bootstrap 5.3.2
- XLSX.js 0.18.5
- Font Awesome 6.4.0
