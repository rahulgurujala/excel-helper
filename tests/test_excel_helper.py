import pytest
import os
from excel_helper import ExcelHelper


@pytest.fixture
def excel_file():
    filename = "test_excel.xlsx"
    yield filename
    if os.path.exists(filename):
        os.remove(filename)


@pytest.fixture
def excel_helper(excel_file):
    helper = ExcelHelper(excel_file)
    helper.create_new_workbook()
    return helper


def test_create_and_save_workbook(excel_helper, excel_file):
    excel_helper.save_workbook()
    assert os.path.exists(excel_file)


def test_write_and_read_cell(excel_helper):
    excel_helper.write_cell(1, 1, "Test")
    assert excel_helper.read_cell(1, 1) == "Test"


def test_write_and_read_row(excel_helper):
    test_data = ["A", "B", "C"]
    excel_helper.write_row(1, test_data)
    assert excel_helper.read_row(1) == test_data


def test_write_and_read_column(excel_helper):
    test_data = ["X", "Y", "Z"]
    excel_helper.write_column(1, test_data)
    assert excel_helper.read_column(1) == test_data


def test_write_and_read_range(excel_helper):
    test_data = [["1", "2"], ["3", "4"]]
    excel_helper.write_range(1, 1, test_data)
    assert excel_helper.read_range(1, 1, 2, 2) == test_data


def test_set_and_get_formula(excel_helper):
    formula = "=SUM(A1:A5)"
    excel_helper.set_formula(1, 1, formula)
    assert excel_helper.get_formula(1, 1) == formula


def test_copy_formula(excel_helper):
    original_formula = "=SUM(A1:A5)"
    excel_helper.set_formula(1, 1, original_formula)
    excel_helper.copy_formula(1, 1, 2, 2)
    copied_formula = excel_helper.get_formula(2, 2)
    assert copied_formula == "=SUM(B1:B5)"


def test_sum_range(excel_helper):
    test_data = [[1], [2], [3], [4], [5]]
    excel_helper.write_range(1, 1, test_data)
    excel_helper.sum_range(1, 1, 5, 1, 6, 1)
    assert excel_helper.get_formula(6, 1) == "=SUM(A1:A5)"


def test_average_range(excel_helper):
    test_data = [[1], [2], [3], [4], [5]]
    excel_helper.write_range(1, 1, test_data)
    excel_helper.average_range(1, 1, 5, 1, 6, 1)
    assert excel_helper.get_formula(6, 1) == "=AVERAGE(A1:A5)"


def test_count_range(excel_helper):
    test_data = [[1], [2], [3], [4], [5]]
    excel_helper.write_range(1, 1, test_data)
    excel_helper.count_range(1, 1, 5, 1, 6, 1)
    assert excel_helper.get_formula(6, 1) == "=COUNT(A1:A5)"


def test_if_formula(excel_helper):
    excel_helper.if_formula(1, 1, "True", "False", 2, 1)
    assert excel_helper.get_formula(2, 1) == '=IF(A1, "True", "False")'


def test_vlookup(excel_helper):
    test_data = [["A", 1], ["B", 2], ["C", 3]]
    excel_helper.write_range(1, 1, test_data)
    excel_helper.write_cell(5, 1, "B")
    excel_helper.vlookup(5, 1, 1, 1, 3, 2, 2, 5, 2)
    assert excel_helper.get_formula(5, 2) == "=VLOOKUP(A5, A1:B3, 2, FALSE)"


def test_select_sheet(excel_helper):
    excel_helper.workbook.create_sheet("TestSheet")
    excel_helper.select_sheet("TestSheet")
    assert excel_helper.active_sheet.title == "TestSheet"


def test_apply_style(excel_helper):
    style = {"font": {"bold": True, "color": "FF0000"}}
    excel_helper.write_cell(1, 1, "Styled Cell")
    excel_helper.apply_style(1, 1, style)
    cell = excel_helper.active_sheet.cell(1, 1)
    assert cell.font.bold
    assert cell.font.color.rgb == "FF0000"


def test_auto_fit_columns(excel_helper):
    test_data = [["Short", "A very long column header"]]
    excel_helper.write_range(1, 1, test_data)
    excel_helper.auto_fit_columns()
    assert (
        excel_helper.active_sheet.column_dimensions["A"].width
        < excel_helper.active_sheet.column_dimensions["B"].width
    )
