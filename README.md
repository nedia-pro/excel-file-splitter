# Excel File Splitter

**Excel File Splitter** is a Python tool that splits large Excel files into smaller, manageable parts. The tool ensures that the headers are preserved, and custom formatting is applied to make the files more organized and readable.

## Features

- Splits large Excel files into smaller chunks based on a predefined row limit.
- Preserves header formatting across all split files.
- Customizable formatting for headers (bold, background color, borders, etc.).
- Supports large datasets (up to 10,000 rows).

## Requirements

- Python 3.x
- `pandas` for data manipulation.
- `openpyxl` for Excel file handling and formatting.

## Installation

To install the necessary libraries, run:

```bash
pip install pandas openpyxl
