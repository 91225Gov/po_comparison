"""
Retail Analytics Manager - Excel File Comparison Application
Upload two Excel files, run verification, and view column-field wise differences.
"""

import streamlit as st
import pandas as pd
from io import BytesIO

from excel_compare import compare_excel_files, ComparisonResult, CellDifference, KEY_COLUMN

st.set_page_config(
    page_title="Retail Analytics – Excel Compare",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom styling for a professional retail analytics look
st.markdown("""
<style>
    .main-header {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1e3a5f;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        color: #5a6c7d;
        font-size: 1rem;
        margin-bottom: 2rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%);
        padding: 1rem 1.25rem;
        border-radius: 8px;
        border-left: 4px solid #3b82f6;
        margin: 0.5rem 0;
    }
    .diff-row {
        padding: 0.5rem;
        border-radius: 4px;
    }
    .stButton > button {
        font-weight: 600;
        border-radius: 8px;
    }
</style>
""", unsafe_allow_html=True)


def load_excel(uploaded_file) -> tuple[dict[str, pd.DataFrame], list[str]]:
    """Load all sheets from an Excel file. Returns dict of sheet_name -> DataFrame and list of sheet names."""
    xl = pd.ExcelFile(uploaded_file)
    sheets = {}
    for name in xl.sheet_names:
        sheets[name] = pd.read_excel(xl, sheet_name=name, header=0)
    return sheets, xl.sheet_names


def run_comparison(
    file1_sheets: dict[str, pd.DataFrame],
    file2_sheets: dict[str, pd.DataFrame],
    sheet1_name: str,
    sheet2_name: str,
    key_column: str = KEY_COLUMN,
) -> ComparisonResult:
    df1 = file1_sheets[sheet1_name]
    df2 = file2_sheets[sheet2_name]
    return compare_excel_files(df1, df2, sheet1_name, sheet2_name, key_column=key_column)


def result_to_dataframe(result: ComparisonResult) -> pd.DataFrame:
    """Convert list of CellDifference to DataFrame for display/export."""
    key_col = result.key_column or KEY_COLUMN
    rows = [
        {
            key_col: d.key_value,
            "Excel Row": d.excel_row,
            "Column": d.column,
            "File 1 Value": d.value_file1,
            "File 2 Value": d.value_file2,
        }
        for d in result.differences
    ]
    return pd.DataFrame(rows)


def main():
    st.markdown('<p class="main-header">📊 Retail Analytics Manager – Excel Comparison</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="sub-header">Compare two Excel files by <strong>Purchasing document number</strong>. For each document in File 1, the matching row in File 2 is found and all columns are compared. Upload files, then run verification.</p>',
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("File 1")
        file1 = st.file_uploader("Choose first Excel file", type=["xlsx", "xls"], key="file1")

    with col2:
        st.subheader("File 2")
        file2 = st.file_uploader("Choose second Excel file", type=["xlsx", "xls"], key="file2")

    file1_sheets = file2_sheets = None
    sheet_names1 = sheet_names2 = []

    if file1:
        try:
            file1_sheets, sheet_names1 = load_excel(file1)
            st.caption(f"File 1: **{file1.name}** — {len(sheet_names1)} sheet(s)")
        except Exception as e:
            st.error(f"Could not read File 1: {e}")

    if file2:
        try:
            file2_sheets, sheet_names2 = load_excel(file2)
            st.caption(f"File 2: **{file2.name}** — {len(sheet_names2)} sheet(s)")
        except Exception as e:
            st.error(f"Could not read File 2: {e}")

    sheet1_name = sheet2_name = None
    if file1_sheets and sheet_names1:
        sheet1_name = st.sidebar.selectbox("Sheet in File 1", sheet_names1, key="sheet1")
    if file2_sheets and sheet_names2:
        sheet2_name = st.sidebar.selectbox("Sheet in File 2", sheet_names2, key="sheet2")

    st.sidebar.markdown("---")
    st.sidebar.caption(f"Unique key: **{KEY_COLUMN}**")
    execute = st.sidebar.button("▶ Run comparison", type="primary", use_container_width=True)

    if execute:
        if not file1 or not file2:
            st.warning("Please upload both Excel files before running the comparison.")
        elif not file1_sheets or not file2_sheets:
            st.warning("Could not load one or both files. Check file format.")
        else:
            with st.spinner("Running verification..."):
                result = run_comparison(
                    file1_sheets, file2_sheets,
                    sheet1_name or sheet_names1[0],
                    sheet2_name or sheet_names2[0],
                    key_column=KEY_COLUMN,
                )

            if result.error:
                st.error(result.error)
                return

            st.success("Verification complete.")
            st.markdown("---")
            st.subheader("Summary")

            summary = result.summary_dict()
            cols = st.columns(3)
            for idx, (k, v) in enumerate(summary.items()):
                with cols[idx % 3]:
                    st.metric(k, v)

            st.markdown("---")
            st.subheader("Columns used for comparison")
            st.caption(f"Only these columns are compared for each {KEY_COLUMN}.")
            if result.columns_compared:
                st.code(", ".join(result.columns_compared))
            if result.requested_columns_missing_in_file1:
                st.warning(f"Requested columns not found in File 1: {', '.join(result.requested_columns_missing_in_file1)}")
            if result.requested_columns_missing_in_file2:
                st.warning(f"Requested columns not found in File 2: {', '.join(result.requested_columns_missing_in_file2)}")

            st.markdown("---")
            st.subheader("Areas of difference")

            if result.keys_only_in_file1:
                st.markdown(f"**{KEY_COLUMN} only in File 1 (missing in File 2):**")
                st.code(", ".join(str(k) for k in result.keys_only_in_file1))
            if result.keys_only_in_file2:
                st.markdown(f"**{KEY_COLUMN} only in File 2 (not in File 1):**")
                st.code(", ".join(str(k) for k in result.keys_only_in_file2))

            if result.columns_only_in_file1:
                st.markdown("**Columns only in File 1:**")
                st.code(", ".join(result.columns_only_in_file1))
            if result.columns_only_in_file2:
                st.markdown("**Columns only in File 2:**")
                st.code(", ".join(result.columns_only_in_file2))

            if result.differences:
                st.markdown(f"**Cell-by-cell differences (by {KEY_COLUMN}, row, column, File 1 value, File 2 value):**")
                diff_df = result_to_dataframe(result)
                st.dataframe(diff_df, use_container_width=True, height=400)

                buf = BytesIO()
                diff_df.to_excel(buf, index=False, sheet_name="Differences")
                buf.seek(0)
                st.download_button(
                    "Download differences as Excel",
                    data=buf,
                    file_name="excel_comparison_differences.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                if result.common_columns:
                    st.info(f"No cell differences found. For every {KEY_COLUMN} in File 1, the matching row in File 2 has the same values in common columns.")
                else:
                    st.info("No common columns to compare.")

    else:
        st.info("Upload two Excel files and click **Run comparison** to see differences.")


if __name__ == "__main__":
    main()
