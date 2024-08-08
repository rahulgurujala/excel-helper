# ExcelHelper

ExcelHelper is a Python library that simplifies Excel manipulation using the openpyxl library. It provides an easy-to-use interface for common Excel operations, including working with formulas.

## Installation

You can install ExcelHelper using pip:

```
pip install excel-helper
```

## Usage

Here's a quick example of how to use ExcelHelper:

```python
from excel_helper import ExcelHelper

# Create a new Excel file
excel = ExcelHelper("example.xlsx")
excel.create_new_workbook()

# Write some data
excel.write_range(1, 1, [
    ["Product", "Quantity", "Price"],
    ["Apple", 10, 0.5],
    ["Banana", 15, 0.3],
    ["Orange", 8, 0.7]
])

# Calculate total for each product
excel.set_formula(2, 4, "=B2*C2")
excel.set_formula(3, 4, "=B3*C3")
excel.set_formula(4, 4, "=B4*C4")

# Calculate sum of quantities and total price
excel.sum_range(2, 2, 4, 2, 5, 2)  # Sum of quantities
excel.sum_range(2, 4, 4, 4, 5, 4)  # Sum of totals

# Calculate average price
excel.average_range(2, 3, 4, 3, 5, 3)

# Use IF formula
excel.if_formula(5, 4, "High Sales", "Low Sales", 5, 5)

# Use VLOOKUP
excel.write_cell(7, 1, "Banana")
excel.vlookup(7, 1, 1, 1, 4, 3, 3, 7, 2)

excel.auto_fit_columns()
excel.save_workbook()
```

## Features

- Open, create, and save workbooks
- Select sheets
- Read and write individual cells
- Read and write rows and columns
- Read and write ranges of cells
- Apply styles to cells
- Auto-fit column widths
- Work with formulas (SUM, AVERAGE, COUNT, IF, VLOOKUP)

## API Reference

### ExcelHelper(filename)

Create a new ExcelHelper instance.

- `filename`: The name of the Excel file to work with.

### Methods

- `open_workbook()`: Open the Excel workbook.
- `save_workbook()`: Save the Excel workbook.
- `create_new_workbook()`: Create a new Excel workbook.
- `select_sheet(sheet_name)`: Select a sheet by name.
- `write_cell(row, col, value)`: Write a value to a specific cell.
- `read_cell(row, col)`: Read the value from a specific cell.
- `write_row(row, data)`: Write a list of values to a row.
- `read_row(row)`: Read all values from a row.
- `write_column(col, data)`: Write a list of values to a column.
- `read_column(col)`: Read all values from a column.
- `write_range(start_row, start_col, data)`: Write a 2D list of values to a range of cells.
- `read_range(start_row, start_col, end_row, end_col)`: Read a range of cells and return a 2D list of values.
- `apply_style(row, col, style)`: Apply a style to a specific cell.
- `auto_fit_columns()`: Auto-fit all columns in the active sheet.
- `set_formula(row, col, formula)`: Set a formula in a specific cell.
- `get_formula(row, col)`: Get the formula from a specific cell.
- `copy_formula(from_row, from_col, to_row, to_col)`: Copy a formula from one cell to another, adjusting cell references.
- `sum_range(start_row, start_col, end_row, end_col, result_row, result_col)`: Sum a range of cells and put the result in another cell.
- `average_range(start_row, start_col, end_row, end_col, result_row, result_col)`: Calculate the average of a range of cells.
- `count_range(start_row, start_col, end_row, end_col, result_row, result_col)`: Count non-empty cells in a range.
- `if_formula(condition_row, condition_col, true_value, false_value, result_row, result_col)`: Set an IF formula in a specific cell.
- `vlookup(lookup_value_row, lookup_value_col, table_start_row, table_start_col, table_end_row, table_end_col, col_index, result_row, result_col)`: Set a VLOOKUP formula in a specific cell.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Support

If you encounter any problems or have any questions, please open an issue on the GitHub repository.