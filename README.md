# FlexiMatch-XL

A flexible Python utility to compare values from an Excel dataset against a list of values in a CSV file and export the results into structured Excel outputs.

This tool is generic and can be used for any type of value matching.

---

## Overview

The scripts read data from an Excel file and compare extracted values against a reference list from a CSV file. The output is written into Excel files.

Two modes are available:

- sprt.py → Categorized output (multiple sheets with summary)
- inonesheet.py → Single-sheet output with match flag

---

## Input Format

### Excel File

- Contains at least one sheet
- Target sheet includes columns with values to be checked
- Each column is processed independently
- Cells may contain:
  - Single value
  - Multiple values separated by spaces

### CSV File

- One value per line
- Extra spaces and trailing commas are automatically cleaned

---

## ⚠️ What You Need to Replace Before Running

Open the script and update the following lines based on your files:

### 1. Excel File Path

Replace:
```
pd.read_excel('compounds.xlsx', sheet_name='Common names', header=0)
```

With:
```
pd.read_excel('your_file.xlsx', sheet_name='your_sheet_name', header=0)
```

---

### 2. CSV File Path

Replace:
```
with open('data.csv', 'r') as f:
```

With:
```
with open('your_file.csv', 'r') as f:
```

---

### 3. Output File Name (Optional)

Replace:
```
compound_matches.xlsx
```
or
```
matches.xlsx
```

With:
```
output.xlsx
```

---

## Dependencies

```
pip install pandas openpyxl
```

---

## Usage

Run scripts from the project directory.

### Run categorized version

```
python sprt.py
```

### Run simple version

```
python inonesheet.py
```

---

## Output

### sprt.py

Creates an Excel file with:

- PRESENT Matches → Values found in CSV
- NOT PRESENT → Values not found in CSV
- Summary → Total processed values and counts

---

### inonesheet.py

Creates an Excel file with:

- Single sheet containing:
  - Column name
  - Extracted value
  - Match status (Yes/No)

---

## Processing Flow

1. Load Excel data  
2. Load CSV data  
3. Clean values (remove spaces, trailing commas)  
4. Iterate through columns and cells  
5. Split multiple values where applicable  
6. Compare each value with CSV dataset  
7. Store results  
8. Export to Excel  

---

## Customization

- Update file paths inside scripts
- Modify sheet name as needed
- Change delimiter if values are not space-separated
- Extend logic for case-insensitive matching if required

---

## Project Structure

```
.
├── sprt.py
├── inonesheet.py
└── README.md
```

---

## Notes

- Matching is case-sensitive by default  
- Ensure correct sheet name is used in the script  
- Empty or null values are ignored  
- Works efficiently with large datasets using pandas  

---

## License

Free to use and modify.
