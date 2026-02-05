# Retail Analytics Manager – Excel File Comparison

Compare two Excel files **column-field wise across all rows**, then review a clear list of differences.

## Features

- **Upload two files**: Choose any two Excel files (`.xlsx` or `.xls`).
- **Sheet selection**: If a workbook has multiple sheets, pick which sheet to use from each file.
- **Run verification**: One-click comparison after both files are uploaded.
- **Summary**: Row counts, common columns, columns only in one file, cells compared, total differences, match %.
- **Differences list**: Every cell difference with Excel row, column name, value in File 1, value in File 2.
- **Export**: Download the differences report as an Excel file.

## How to run

1. **Create a virtual environment (recommended):**
   ```bash
   cd file-comparison
   python3 -m venv venv
   source venv/bin/activate   # On Windows: venv\Scripts\activate
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Start the app:**
   ```bash
   streamlit run app.py
   ```

4. Open the URL shown in the terminal (usually `http://localhost:8501`).

5. **Use the app:**
   - Upload **File 1** and **File 2** in the two upload areas.
   - Optionally select which sheet to use for each file in the sidebar.
   - Click **Run comparison** in the sidebar.
   - Review the summary and the table of differences; download the report if needed.

## Comparison logic

- Rows are aligned by **position** (row 1 in File 1 vs row 1 in File 2, etc.).
- Only **columns that exist in both files** are compared cell by cell.
- Columns that appear in only one file are listed separately (not cell-compared).
- Empty/missing rows in the shorter file are reported as differences (e.g. “no row” in the other file).
- Numeric and text values are normalized for comparison (e.g. spaces trimmed, NaN treated as empty).

## Deploy on Streamlit Community Cloud

Deploy from the **Streamlit Cloud website** (don’t use the “Deploy” button inside the running app).

1. **Push your code to GitHub** (if not already):
   ```bash
   cd file-comparison   # or your repo folder
   git add .
   git commit -m "Your message"
   git push -u origin main
   ```

2. **Open Streamlit Community Cloud:**  
   Go to [share.streamlit.io](https://share.streamlit.io) and sign in with **GitHub**.

3. **Create a new app:**
   - Click **“New app”**.
   - **Repository:** `91225Gov/po_comparison` (or select it from the list).
   - **Branch:** `main`.
   - **Main file path:** `app.py`.
   - Click **“Deploy!”**.

4. Wait for the build to finish. Your app will be available at a URL like `https://po-comparison-xxxx.streamlit.app`.

If you see *“The app’s code is not connected to a remote GitHub repository”*, it usually means you’re using the in-app **Deploy** button. Use the steps above on [share.streamlit.io](https://share.streamlit.io) instead, and select your GitHub repo there.

## Requirements

- Python 3.9+
- See `requirements.txt` for Python packages (Streamlit, pandas, openpyxl, xlrd).
