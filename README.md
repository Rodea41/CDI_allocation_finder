<img src="https://github.com/Rodea41/CDI_allocation_finder/blob/main/butter.webp" width="1000" height="500" />


# CDI Inventory Reader

## Overview

The CDI Inventory Reader is a Python tool designed to automate the process of matching CDI allocation PDFs with inventory data. This utility extracts pallet IDs and quantity information from CDI order manifests, cross-references them with inventory data, and generates formatted Excel reports for each order.

## Features

- **PDF Processing**: Automatically reads and extracts data from CDI allocation PDFs
- **Inventory Matching**: Cross-references pallet IDs with inventory data
- **Lot Tracking**: Identifies all pallets belonging to the same lot numbers
- **Report Generation**: Creates formatted Excel reports with color-coded highlighting
- **Batch Processing**: Handles multiple allocation PDFs in a single run

## Requirements

The script requires the following Python libraries:
- PyPDF2
- pandas
- openpyxl
- csv
- re
- os

You can install these dependencies using pip:

```bash
pip install PyPDF2 pandas openpyxl
```

## Configuration

Before running the script, you need to configure the following directory paths in the code:

1. **Inventory Directory**: The location where your inventory CSV file is stored
   ```python
   inventory_directory = "C:\\Users\\crodea\\Desktop\\OneDrive\\OneDrive - US Cold Storage\\Python\\CDI_allocations_project"
   ```

2. **Inventory File**: The name of your inventory CSV file
   ```python
   inventory_file = "inventory.csv"
   ```

3. **Allocation PDFs Directory**: The folder containing CDI allocation PDFs
   ```python
   path = "C:\\Users\\crodea\\Desktop\\test_allocations"
   ```

## How It Works

1. **File Preparation**: The script renames all PDFs in the allocation directory to a standard format (CDI0.pdf, CDI1.pdf, etc.)
2. **Data Extraction**: For each PDF:
   - Extracts the order number (SO00xxxxxxx)
   - Identifies all pallet IDs (format: xxxx-xxx-xx-xxx-x)
   - Extracts requested quantities
3. **Inventory Matching**: Looks up each pallet ID in the inventory CSV
4. **Lot Analysis**: Identifies all pallets from the same lots
5. **Report Generation**: Creates a CSV and formatted Excel file for each order

## Output

For each processed order, the script generates:
1. A CSV file named with the order number
2. An Excel file with the same name, featuring:
   - Color-coded headers
   - Bold formatting for important information
   - Highlighted license numbers
   - Organized sections for matched pallets and lot information

## Usage

Simply run the script:

```python
python CDI_inventory_reader.py
```

The script will process all PDFs in the configured allocation directory and generate reports in the same location.

## Key Functions

- `getOrderNumber()`: Extracts the order number from PDF text
- `getPalletId()`: Identifies pallet IDs in the allocation
- `getQTY()`: Extracts requested quantities
- `read_from_inventory_csv()`: Matches pallet IDs with inventory data
- `get_entire_lot()`: Identifies all lots in the order
- `read_lots_from_csv()`: Retrieves all pallets belonging to identified lots
- `format_and_style()`: Applies formatting to the Excel output
