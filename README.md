# TKL Value Classifier

## Overview
This project processes Excel files within a specified input directory and extracts unique values from specific columns. These values are then appended to designated text files in an output directory. The script ensures no duplicates are added and is designed for batch processing of subdirectories containing Excel files.

## Features
- Extracts unique values from specific columns (Ex. columns 5, 6, 10, and 12).
- Appends extracted values to respective text files:
  - `projectTypeTKLValues.txt`
  - `documentTypeTKLValues.txt`
  - `metadataReviewedTKLValues.txt`
- Handles batch processing of subdirectories containing Excel files with name pattern Index<*name*>.xlsx.
- Prevents duplicate entries by maintaining existing unique values in the text files.

## Why This Project is Useful
This tool automates the tedious task of extracting and consolidating unique values from Excel files into categorized text files. It is especially useful for projects that require standardization and organization of large datasets.

## How to Use
### Prerequisites
- Python 3.6 or above
- The `openpyxl` library (`pip install openpyxl`)

### Usage
1. Save the script as `extractor.py`.
2. Run the script using the following command:
   ```bash
   python extractor.py <input_directory> <output_directory>
