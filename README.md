### **Data Extraction and Consolidation Tool**

This Python script is designed to streamline the process of extracting and consolidating email and phone number data from various file formats, including Excel and CSV. It allows users to handle multiple files or an entire folder, with flexibility to accommodate different column structures.

#### **Features:**

- **File Handling**: Supports extraction from both Excel (`.xlsx`, `.xls`) and CSV (`.csv`) files.
- **Flexible Input**: Users can choose to input multiple files or a folder containing the files.
- **Column Specification**: Provides the ability to specify which columns contain email and phone number data for both Excel and CSV files.
- **Robust Parsing**: Handles both multi-column and single-column CSV formats. It ensures that data is correctly parsed and formatted, even with varying row structures.
- **Error Handling**: Includes error handling for unexpected formats and issues during processing.
- **Output**: Consolidates extracted data into a single DataFrame and saves the result to a new Excel file.

#### **Usage:**

1. **File Input**: Users are prompted to specify whether they are providing multiple files or a folder. For multiple files, paths are entered separated by commas. For a folder, the folder path is provided.
   
2. **Column Identification**: For each file, users specify the columns where email and phone number data are located. The script adjusts according to whether the input is in a single-column CSV format or a multi-column format.

3. **Data Extraction**: The script extracts email and phone number data from each file, handling various formats and ensuring data integrity.

4. **Data Consolidation**: Extracted data from all files are combined into a single DataFrame.

5. **Output File**: The final consolidated data is saved to an Excel file at a user-specified location.

#### **Requirements:**

- Python 3.x
- Pandas library (`pip install pandas`)
- Openpyxl library for Excel file handling (`pip install openpyxl`)

#### **Example:**

```python
python extract_and_consolidate.py
```

**Note:** Adjust the column indices or names as per your file structure for accurate extraction.
