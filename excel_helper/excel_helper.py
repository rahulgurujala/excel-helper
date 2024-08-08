from typing import Any, Dict, List

import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter


class ExcelHelper:
    def __init__(self, filename: str):
        self.filename = filename
        self.workbook = None
        self.active_sheet = None

    def open_workbook(self):
        """Open the Excel workbook."""
        self.workbook = openpyxl.load_workbook(self.filename)
        self.active_sheet = self.workbook.active

    def save_workbook(self):
        """Save the Excel workbook."""
        self.workbook.save(self.filename)

    def create_new_workbook(self):
        """Create a new Excel workbook."""
        self.workbook = openpyxl.Workbook()
        self.active_sheet = self.workbook.active

    def select_sheet(self, sheet_name: str):
        """Select a sheet by name."""
        if sheet_name in self.workbook.sheetnames:
            self.active_sheet = self.workbook[sheet_name]
        else:
            raise ValueError(f"Sheet '{sheet_name}' not found in the workbook.")

    def write_cell(self, row: int, col: int, value: Any):
        """Write a value to a specific cell."""
        self.active_sheet.cell(row=row, column=col, value=value)

    def read_cell(self, row: int, col: int) -> Any:
        """Read the value from a specific cell."""
        return self.active_sheet.cell(row=row, column=col).value

    def write_row(self, row: int, data: List[Any]):
        """Write a list of values to a row."""
        for col, value in enumerate(data, start=1):
            self.write_cell(row, col, value)

    def read_row(self, row: int) -> List[Any]:
        """Read all values from a row."""
        return [cell.value for cell in self.active_sheet[row]]

    def write_column(self, col: int, data: List[Any]):
        """Write a list of values to a column."""
        for row, value in enumerate(data, start=1):
            self.write_cell(row, col, value)

    def read_column(self, col: int) -> List[Any]:
        """Read all values from a column."""
        return [cell.value for cell in self.active_sheet[get_column_letter(col)]]

    def write_range(self, start_row: int, start_col: int, data: List[List[Any]]):
        """Write a 2D list of values to a range of cells."""
        for row_offset, row_data in enumerate(data):
            for col_offset, value in enumerate(row_data):
                self.write_cell(start_row + row_offset, start_col + col_offset, value)

    def read_range(
        self, start_row: int, start_col: int, end_row: int, end_col: int
    ) -> List[List[Any]]:
        """Read a range of cells and return a 2D list of values."""
        return [
            [cell.value for cell in row]
            for row in self.active_sheet.iter_rows(
                min_row=start_row, min_col=start_col, max_row=end_row, max_col=end_col
            )
        ]

    def apply_style(self, row: int, col: int, style: Dict[str, Any]):
        """Apply a style to a specific cell."""
        cell = self.active_sheet.cell(row=row, column=col)
        for key, value in style.items():
            setattr(cell, key, value)

    def auto_fit_columns(self):
        """Auto-fit all columns in the active sheet."""
        for column in self.active_sheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except Exception:
                    pass
            adjusted_width = max_length + 2
            self.active_sheet.column_dimensions[column_letter].width = adjusted_width

    def set_formula(self, row: int, col: int, formula: str):
        """Set a formula in a specific cell."""
        self.active_sheet.cell(row=row, column=col, value=formula)

    def get_formula(self, row: int, col: int) -> str:
        """Get the formula from a specific cell."""
        return (
            self.active_sheet.cell(row=row, column=col).data_type == "f"
            and self.active_sheet.cell(row=row, column=col).value
        )

    def copy_formula(self, from_row: int, from_col: int, to_row: int, to_col: int):
        """Copy a formula from one cell to another, adjusting cell references."""
        source_cell = self.active_sheet.cell(row=from_row, column=from_col)
        target_cell = self.active_sheet.cell(row=to_row, column=to_col)

        if source_cell.data_type == "f":
            translated_formula = Translator(
                source_cell.value, origin=source_cell.coordinate
            ).translate_formula(target_cell.coordinate)
            target_cell.value = translated_formula

    def sum_range(
        self,
        start_row: int,
        start_col: int,
        end_row: int,
        end_col: int,
        result_row: int,
        result_col: int,
    ):
        """Sum a range of cells and put the result in another cell."""
        start_cell = get_column_letter(start_col) + str(start_row)
        end_cell = get_column_letter(end_col) + str(end_row)
        formula = f"=SUM({start_cell}:{end_cell})"
        self.set_formula(result_row, result_col, formula)

    def average_range(
        self,
        start_row: int,
        start_col: int,
        end_row: int,
        end_col: int,
        result_row: int,
        result_col: int,
    ):
        """Calculate the average of a range of cells and put the result in another cell."""
        start_cell = get_column_letter(start_col) + str(start_row)
        end_cell = get_column_letter(end_col) + str(end_row)
        formula = f"=AVERAGE({start_cell}:{end_cell})"
        self.set_formula(result_row, result_col, formula)

    def count_range(
        self,
        start_row: int,
        start_col: int,
        end_row: int,
        end_col: int,
        result_row: int,
        result_col: int,
    ):
        """Count non-empty cells in a range and put the result in another cell."""
        start_cell = get_column_letter(start_col) + str(start_row)
        end_cell = get_column_letter(end_col) + str(end_row)
        formula = f"=COUNT({start_cell}:{end_cell})"
        self.set_formula(result_row, result_col, formula)

    def if_formula(
        self,
        condition_row: int,
        condition_col: int,
        true_value: Any,
        false_value: Any,
        result_row: int,
        result_col: int,
    ):
        """Set an IF formula in a specific cell."""
        condition_cell = get_column_letter(condition_col) + str(condition_row)
        formula = f'=IF({condition_cell}, "{true_value}", "{false_value}")'
        self.set_formula(result_row, result_col, formula)

    def vlookup(
        self,
        lookup_value_row: int,
        lookup_value_col: int,
        table_start_row: int,
        table_start_col: int,
        table_end_row: int,
        table_end_col: int,
        col_index: int,
        result_row: int,
        result_col: int,
    ):
        """Set a VLOOKUP formula in a specific cell."""
        lookup_value = get_column_letter(lookup_value_col) + str(lookup_value_row)
        table_range = f"{get_column_letter(table_start_col)}{table_start_row}:{get_column_letter(table_end_col)}{table_end_row}"
        formula = f"=VLOOKUP({lookup_value}, {table_range}, {col_index}, FALSE)"
        self.set_formula(result_row, result_col, formula)


# Example usage
if __name__ == "__main__":
    excel = ExcelHelper("example.xlsx")
    excel.create_new_workbook()

    # Write some sample data
    excel.write_range(
        1,
        1,
        [
            ["Product", "Quantity", "Price"],
            ["Apple", 10, 0.5],
            ["Banana", 15, 0.3],
            ["Orange", 8, 0.7],
        ],
    )

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
