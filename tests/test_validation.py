
import pandas as pd
from app.main import run_validation


def test_valid_data_passes():
    df = pd.DataFrame({"SKU": ["A1", "B2"], "Quantity": [2, 3]})
    result = run_validation(df, required_columns=["SKU", "Quantity"], unique_columns=["SKU"])
    assert result["status"] == "PASS"
    assert result["data_rows_checked"] == 2
    assert result["error_count"] == 0


def test_missing_required_column_fails():
    df = pd.DataFrame({"SKU": ["A1"]})
    result = run_validation(df, required_columns=["SKU", "Quantity"])
    assert result["status"] == "FAIL"
    assert any(i["issue"] == "Required column missing" for i in result["issues"])


def test_blank_mandatory_cell_fails():
    df = pd.DataFrame({"SKU": ["A1", ""]})
    result = run_validation(df, required_columns=["SKU"])
    assert result["status"] == "FAIL"
    assert any(i["issue"] == "Required value is blank" and i["row"] == 3 for i in result["issues"])


def test_duplicate_identifier_fails():
    df = pd.DataFrame({"Order ID": ["X1", "X1"]})
    result = run_validation(df, unique_columns=["Order ID"])
    assert result["status"] == "FAIL"
    assert sum(i["category"] == "duplicate_identifier" for i in result["issues"]) == 2


def test_duplicate_row_is_warning():
    df = pd.DataFrame({"SKU": ["A1", "A1"], "Quantity": [1, 1]})
    result = run_validation(df)
    assert result["warning_count"] == 1
    assert result["status"] == "PASS WITH WARNINGS"


def test_numeric_heuristic_warns_not_errors():
    df = pd.DataFrame({"Quantity": ["two"]})
    result = run_validation(df)
    assert result["warning_count"] == 1
    assert result["status"] == "PASS WITH WARNINGS"
