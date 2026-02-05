"""
Retail Analytics - Excel File Comparison Engine
Compares two Excel files column-field wise across all rows and reports differences.
"""

import pandas as pd
from dataclasses import dataclass, field

KEY_COLUMN = "Purchasing document number"

# Columns to compare for each Purchasing document number (only these are compared)
COMPARE_COLUMNS = [
    "Purchasing document number",
    "0GN_VENDOR",
    "Purchasing Document Type",
    "Puchasing document category",
    "Source System for R/3 Entity",
    "Invoicing Party",
    "Logical System Backend",
    "Status of purchasing document",
    "Purchasing Organization",
    "Purchasing Group",
    "Source system ID",
    "Supplying Plant",
    "0VENDOR",
    "Purchase Document Date",
    "Date on which the purchasing document wa",
    "Fiscal year / period",
    "Fiscal year variant",
    "Fiscal year",
    "Posting Date in the Document",
    "Update Date",
    "Validity period end",
    "Validity Period Start",
    "Calendar Day Number",
    "Calender Week - Saturday",
    "Calender Week - Sunday",
    "Calendar Week Number",
    "Document Currency",
    "Purchase Order Currency",
    "Company Code",
    "Name of person who created the object",
    "Date on which the record was created",
    "Transaction for purchasing document",
    "Logical System",
    "Ordering address",
    "Vendor to whom partner roles are assigne",
    "Goods Supplier",
    "Supplying vendor",
    "Supplying plant to which partner roles h",
    "Released Date",
    "Released By",
    "PO release indicator",
    "PO release status",
    "Number of deliveries",
    "Exchange rate for pricing and statistics",
    "No. of invoices",
    "Counter for documents",
    "Re-Release Count",
    "Total value at the time of Release",
    "Total value at the time of ReleaseTotal value at the time of Release",  # in case header is duplicated in Excel
]


@dataclass
class CellDifference:
    """Single cell difference between two files."""
    row_index: int
    column: str
    value_file1: object
    value_file2: object
    excel_row: int  # 1-based for Excel display (header = 1)
    key_value: object = None  # Purchasing document number for this row


@dataclass
class ComparisonResult:
    """Result of comparing two Excel DataFrames."""
    differences: list[CellDifference] = field(default_factory=list)
    columns_only_in_file1: list[str] = field(default_factory=list)
    columns_only_in_file2: list[str] = field(default_factory=list)
    common_columns: list[str] = field(default_factory=list)
    rows_file1: int = 0
    rows_file2: int = 0
    cells_compared: int = 0
    total_differences: int = 0
    sheet_name_file1: str = ""
    sheet_name_file2: str = ""
    key_column: str = KEY_COLUMN
    keys_only_in_file1: list = field(default_factory=list)
    keys_only_in_file2: list = field(default_factory=list)
    error: str | None = None
    columns_compared: list[str] = field(default_factory=list)  # columns actually compared
    requested_columns_missing_in_file1: list[str] = field(default_factory=list)
    requested_columns_missing_in_file2: list[str] = field(default_factory=list)

    @property
    def match_percentage(self) -> float:
        if self.cells_compared == 0:
            return 100.0
        return round(100.0 * (1 - self.total_differences / self.cells_compared), 2)

    def summary_dict(self) -> dict:
        return {
            "Rows in File 1": self.rows_file1,
            "Rows in File 2": self.rows_file2,
            "Common columns": len(self.common_columns),
            "Columns only in File 1": len(self.columns_only_in_file1),
            "Columns only in File 2": len(self.columns_only_in_file2),
            "Cells compared": self.cells_compared,
            "Total differences": self.total_differences,
            "Match %": f"{self.match_percentage}%",
            "Keys only in File 1": len(self.keys_only_in_file1),
            "Keys only in File 2": len(self.keys_only_in_file2),
            "Columns compared (requested)": len(self.columns_compared),
        }


def _normalize_value(val) -> str:
    """Normalize for comparison (handle NaN, types)."""
    if pd.isna(val):
        return ""
    if isinstance(val, (int, float)):
        return str(val).strip()
    return str(val).strip() if val is not None else ""


def _build_key_to_row_index(df: pd.DataFrame, key_column: str) -> dict[str, int]:
    """Build mapping from normalized key value to first row index in df."""
    if key_column not in df.columns:
        return {}
    out = {}
    for i in range(len(df)):
        k = _normalize_value(df[key_column].iloc[i])
        if k not in out:
            out[k] = i
    return out


def compare_excel_files(
    df1: pd.DataFrame,
    df2: pd.DataFrame,
    sheet_name1: str = "Sheet1",
    sheet_name2: str = "Sheet1",
    key_column: str = KEY_COLUMN,
) -> ComparisonResult:
    """
    Compare two DataFrames using a unique key column (e.g. Purchasing document number).
    For each key in File 1, find the matching row in File 2 and compare all common columns.
    """
    result = ComparisonResult(
        sheet_name_file1=sheet_name1,
        sheet_name_file2=sheet_name2,
        rows_file1=len(df1),
        rows_file2=len(df2),
        key_column=key_column,
    )

    if key_column not in df1.columns:
        result.error = f"Key column '{key_column}' not found in File 1. Available columns: {list(df1.columns)}"
        return result
    if key_column not in df2.columns:
        result.error = f"Key column '{key_column}' not found in File 2. Available columns: {list(df2.columns)}"
        return result

    cols1 = set(df1.columns)
    cols2 = set(df2.columns)
    result.common_columns = sorted(cols1 & cols2)
    result.columns_only_in_file1 = sorted(cols1 - cols2)
    result.columns_only_in_file2 = sorted(cols2 - cols1)

    # Only compare the requested columns that exist in both files
    result.columns_compared = [c for c in COMPARE_COLUMNS if c in cols1 and c in cols2]
    result.requested_columns_missing_in_file1 = [c for c in COMPARE_COLUMNS if c not in cols1]
    result.requested_columns_missing_in_file2 = [c for c in COMPARE_COLUMNS if c not in cols2]
    compare_columns = result.columns_compared

    # Build key -> first row index in File 2
    key_to_row2 = _build_key_to_row_index(df2, key_column)
    keys_in_file1 = set()
    for i in range(len(df1)):
        key_val = df1[key_column].iloc[i]
        key_norm = _normalize_value(key_val)
        keys_in_file1.add(key_norm)

    keys_in_file2 = set(key_to_row2.keys())
    result.keys_only_in_file1 = sorted(keys_in_file1 - keys_in_file2)
    result.keys_only_in_file2 = sorted(keys_in_file2 - keys_in_file1)

    # For each row in File 1, find matching row in File 2 and compare
    for i in range(len(df1)):
        key_val = df1[key_column].iloc[i]
        key_norm = _normalize_value(key_val)
        j = key_to_row2.get(key_norm)

        excel_row = i + 2  # 1-based + header

        if j is None:
            # Key exists in File 1 but not in File 2: report all common columns as diff (missing in File 2)
            for col in compare_columns:
                result.cells_compared += 1
                v1 = df1[col].iloc[i]
                result.differences.append(
                    CellDifference(
                        row_index=i,
                        column=col,
                        value_file1=v1,
                        value_file2="(missing in File 2)",
                        excel_row=excel_row,
                        key_value=key_val,
                    )
                )
            continue

        # Matched row: compare each common column
        for col in compare_columns:
            result.cells_compared += 1
            v1 = df1[col].iloc[i]
            v2 = df2[col].iloc[j]
            n1 = _normalize_value(v1)
            n2 = _normalize_value(v2)
            if n1 != n2:
                result.differences.append(
                    CellDifference(
                        row_index=i,
                        column=col,
                        value_file1=v1,
                        value_file2=v2,
                        excel_row=excel_row,
                        key_value=key_val,
                    )
                )

    result.total_differences = len(result.differences)
    return result
