---

# Excel Data Cleaner

![Python Version](https://img.shields.io/badge/python-3.8%2B-blue)
![Tkinter](https://img.shields.io/badge/tkinter-8.6-blue)

## Overview

**Excel Data Cleaner** is a user-friendly tool designed to simplify the process of cleaning and processing Excel files. It provides a graphical interface for tasks such as removing duplicates, filling or dropping missing values, converting date formats, replacing values, correcting port names using fuzzy matching, converting text to uppercase, and combining import/export sheets.

## Features

- **Remove Duplicates**: Easily remove duplicate rows in your Excel data.
- **Fill Missing Values**: Fill missing values with a specified value.
- **Drop Missing Values**: Drop rows that contain any missing values.
- **Convert Date Formats**: Convert date columns to a specified format.
- **Replace Values**: Replace specific values in any column.
- **Correct Port Names**: Automatically correct port names using fuzzy matching.
- **Convert Text to Uppercase**: Convert all text data to uppercase.
- **Remove Leading/Trailing Spaces**: Strip spaces from text columns.
- **Combine Import/Export Sheets**: Combine data from 'IMPORT' and 'EXPORT' sheets into a single sheet.

## Requirements

- **Python 3.8+**
- **Tkinter** (usually included with Python)
- **Pandas**
- **RapidFuzz**

## Installation

### Option 1: Running from Source

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/yourusername/ExcelDataCleaner.git
   cd ExcelDataCleaner
   ```

2. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the Application**:
   ```bash
   python main.py
   ```

### Option 2: Using the Executable

1. **Download the Executable**:
   - Go to the [Releases](https://github.com/yourusername/ExcelDataCleaner/releases) section of this repository.
   - Download the `main.exe` file.

2. **Run the Executable**:
   - Double-click on `main.exe` to start the application.

## Usage

1. **Upload an Excel File**: Click "Upload Excel File" and select the file you want to clean.
2. **Select a Sheet**: Choose the sheet you want to work with from the dropdown menu.
3. **Choose Cleaning Options**: Click on the buttons to perform various cleaning tasks.
4. **Download Cleaned File**: After cleaning, download your cleaned Excel file by clicking "Download Cleaned File".

## Contributing

Contributions are welcome! Please open an issue or submit a pull request if you have suggestions or improvements.

## Acknowledgements

- **Pandas**: For powerful data manipulation tools.
- **Tkinter**: For the graphical user interface.
- **RapidFuzz**: For fast and accurate string matching.

## Contact

For any questions, feel free to reach out via [GitHub Issues](https://github.com/yourusername/ExcelDataCleaner/issues) or contact me directly at `icube.india@gac.com`.

---
