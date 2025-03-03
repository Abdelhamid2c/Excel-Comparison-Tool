# Excel Comparison Tool

## Description

This script compares two Excel files and highlights the differences. It marks changed cells with a red fill and adds comments detailing the previous and new values.

## Requirements

Ensure you have the following Python libraries installed:

```bash
pip install openpyxl pandas
```

## Functions

### 1. `add_comment(old_value, new_value)`

This function generates a comment containing the old and new values and assigns it to the corresponding Excel cell.

#### Parameters:

- `old_value` (any): The original value from the reference Excel file.
- `new_value` (any): The updated value from the new Excel file.

### 2. `compare_excels(ref_path, new_path, output_path)`

This function compares two Excel files and highlights differences.

#### Parameters:

- `ref_path` (str): Path to the reference (older version) Excel file.
- `new_path` (str): Path to the new version Excel file.
- `output_path` (str): Path to save the output Excel file with highlighted changes.

#### Process:

1. Loads both Excel files as Pandas DataFrames.
2. Compares the data cell by cell.
3. Highlights changed cells in red.
4. Adds a comment specifying the old and new values.
5. Saves the updated file to `output_path`.

## Usage Example

```python
ref_path = 'data/reference_file.xlsx'  # Chemin du fichier de référence
new_path = 'data/new_file.xlsx'        # Chemin du nouveau fichier à comparer ou modifier
output_path = 'data/output_file.xlsx'  # Chemin du fichier de sortie après traitement

compare_excels(ref_path, new_path, output_path)
```

After running the script, the `output_file.xlsx` file will contain highlighted differences and comments indicating the value changes.

## Error Handling

- Prints an error message if the Excel files have different shapes.
- Handles file not found errors.
- Catches and prints any other unexpected errors.

## Notes

- Ensure the Excel files have the same structure (number of rows/columns) before running the comparison.
- The script modifies a copy of the old file and saves it as `output_path`.

## Author

Abdelhamid Chebel



