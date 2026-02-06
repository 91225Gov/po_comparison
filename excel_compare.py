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
]


@dataclass
class CellDifference:
    """Single cell difference between two files."""
    row_index: int
    column: str
    value_file1: object
    value_file2: object
    excel_row: int  # 1-based Excel row in File 1 (header = 1)
    excel_row_file2: int | None = None  # 1-based Excel row in File 2, or None if missing
    key_value: object = None  # Key (e.g. Purchasing document number) for this row


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
    # Crosstab per key: list of {key_value, excel_row_file1, excel_row_file2, rows: [{column, file1, file2, is_difference}]}
    key_crosstabs: list = field(default_factory=list)

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


def _row_key(df: pd.DataFrame, key_columns: list[str], i: int) -> tuple:
    """Build composite key (tuple of normalized values) for row i."""
    return tuple(_normalize_value(df[col].iloc[i]) for col in key_columns)


def _row_key_display(df: pd.DataFrame, key_columns: list[str], i: int) -> str:
    """Display string for key (e.g. 'val1' or 'col1=val1, col2=val2')."""
    parts = [f"{col}={df[col].iloc[i]}" for col in key_columns]
    return ", ".join(parts) if len(parts) > 1 else str(df[key_columns[0]].iloc[i])


def _build_key_to_row_index(df: pd.DataFrame, key_columns: list[str]) -> dict[tuple, int]:
    """Build mapping from composite key (tuple) to first row index in df."""
    missing = [c for c in key_columns if c not in df.columns]
    if missing:
        return {}
    out = {}
    for i in range(len(df)):
        k = _row_key(df, key_columns, i)
        if k not in out:
            out[k] = i
    return out


def compare_excel_files(
    df1: pd.DataFrame,
    df2: pd.DataFrame,
    sheet_name1: str = "Sheet1",
    sheet_name2: str = "Sheet1",
    key_columns: list[str] | None = None,
) -> ComparisonResult:
    """
    Compare two DataFrames using one or more unique key columns.
    For each key in File 1, find the matching row in File 2 and compare all common columns.
    """
    if key_columns is None or len(key_columns) == 0:
        key_columns = [KEY_COLUMN]
    key_column_display = ", ".join(key_columns)

    result = ComparisonResult(
        sheet_name_file1=sheet_name1,
        sheet_name_file2=sheet_name2,
        rows_file1=len(df1),
        rows_file2=len(df2),
        key_column=key_column_display,
    )

    missing1 = [c for c in key_columns if c not in df1.columns]
    if missing1:
        result.error = f"Key column(s) {missing1} not found in File 1. Available: {list(df1.columns)}"
        return result
    missing2 = [c for c in key_columns if c not in df2.columns]
    if missing2:
        result.error = f"Key column(s) {missing2} not found in File 2. Available: {list(df2.columns)}"
        return result

    cols1 = set(df1.columns)
    cols2 = set(df2.columns)
    result.common_columns = sorted(cols1 & cols2)
    result.columns_only_in_file1 = sorted(cols1 - cols2)
    result.columns_only_in_file2 = sorted(cols2 - cols1)

    # Only compare the requested columns that exist in both files; order as in File 1
    requested_in_both = [c for c in COMPARE_COLUMNS if c in cols1 and c in cols2]
    result.columns_compared = [c for c in df1.columns if c in requested_in_both]
    result.requested_columns_missing_in_file1 = [c for c in COMPARE_COLUMNS if c not in cols1]
    result.requested_columns_missing_in_file2 = [c for c in COMPARE_COLUMNS if c not in cols2]
    compare_columns = result.columns_compared

    # Build key -> first row index in File 2
    key_to_row2 = _build_key_to_row_index(df2, key_columns)
    keys_in_file1 = set()
    for i in range(len(df1)):
        keys_in_file1.add(_row_key(df1, key_columns, i))

    keys_in_file2 = set(key_to_row2.keys())
    result.keys_only_in_file1 = sorted(str(k) for k in (keys_in_file1 - keys_in_file2))
    result.keys_only_in_file2 = sorted(str(k) for k in (keys_in_file2 - keys_in_file1))

    # For each row in File 1, find matching row in File 2 and compare
    for i in range(len(df1)):
        key_tuple = _row_key(df1, key_columns, i)
        j = key_to_row2.get(key_tuple)
        key_val = _row_key_display(df1, key_columns, i)

        excel_row_file1 = i + 2  # 1-based Excel row in File 1 (header = 1)
        excel_row_file2_val = (j + 2) if j is not None else None  # 1-based Excel row in File 2

        if j is None:
            # Key exists in File 1 but not in File 2: report all common columns as diff (missing in File 2)
            crosstab_rows = []
            for col in compare_columns:
                result.cells_compared += 1
                v1 = df1[col].iloc[i]
                crosstab_rows.append({"column": col, "file1": v1, "file2": "(missing in File 2)", "is_difference": True})
                result.differences.append(
                    CellDifference(
                        row_index=i,
                        column=col,
                        value_file1=v1,
                        value_file2="(missing in File 2)",
                        excel_row=excel_row_file1,
                        excel_row_file2=excel_row_file2_val,
                        key_value=key_val,
                    )
                )
            result.key_crosstabs.append({
                "key_value": key_val,
                "excel_row_file1": excel_row_file1,
                "excel_row_file2": excel_row_file2_val,
                "rows": crosstab_rows,
            })
            continue

        # Matched row: compare each common column and build crosstab
        crosstab_rows = []
        has_diff = False
        for col in compare_columns:
            result.cells_compared += 1
            v1 = df1[col].iloc[i]
            v2 = df2[col].iloc[j]
            n1 = _normalize_value(v1)
            n2 = _normalize_value(v2)
            is_diff = n1 != n2
            if is_diff:
                has_diff = True
                result.differences.append(
                    CellDifference(
                        row_index=i,
                        column=col,
                        value_file1=v1,
                        value_file2=v2,
                        excel_row=excel_row_file1,
                        excel_row_file2=excel_row_file2_val,
                        key_value=key_val,
                    )
                )
            crosstab_rows.append({"column": col, "file1": v1, "file2": v2, "is_difference": is_diff})
        if has_diff:
            result.key_crosstabs.append({
                "key_value": key_val,
                "excel_row_file1": excel_row_file1,
                "excel_row_file2": excel_row_file2_val,
                "rows": crosstab_rows,
            })

    result.total_differences = len(result.differences)
    return result
