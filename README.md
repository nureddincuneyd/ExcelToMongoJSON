# Excel to JSON Converter

This project is a Python script that converts specific columns from an Excel file into a JSON array. The script reads an Excel file, extracts data from specified columns, and writes the data to a JSON file.

## Requirements

- Python 3.x
- pandas
- openpyxl

## Installation

To install the required packages, you can use pip:

```bash
pip install pandas openpyxl
Usage
To run the script, use the following command:

bash
Kodu kopyala
python main.py <source_excel_file> <start_column> <end_column> [output_file]
<source_excel_file>: The path to the source Excel file.
<start_column>: The starting column (e.g., A).
<end_column>: The ending column (e.g., C).
[output_file]: (Optional) The name of the output JSON file. If not provided, the default is output.json.
Examples
bash
Kodu kopyala
python main.py data.xlsx A C output.json
This command will read the data from columns A to C in data.xlsx and write it to output.json.

bash
Kodu kopyala
python main.py data.xlsx A C
This command will read the data from columns A to C in data.xlsx and write it to output.json (default name).

Notes
The script handles empty cells by replacing them with empty strings in the JSON output.
License
This project is licensed under the MIT License. See the LICENSE file for details.
