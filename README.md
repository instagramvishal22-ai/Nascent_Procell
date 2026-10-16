# Procell — Excel Mapping Quality Checker

Procell checks an Excel source-to-target mapping worksheet for:

1. **Duplicate mapping rows** where `Target Table`, `Target Column`, `Pseudo Logic`, and `Transpose Rule` match.
2. **Missing mandatory target columns** for each populated `Source Table` + `Target Table` pair.

Mandatory `Target Column` values: `USUBJID`, `SUBJID`, `SPID`, `INVID`, `VISIT`, `VISITNUM`, `DCMDATE`, `GRPID`, `STUDYID`, `DOMAIN`, and `CVISIT`.

## Run on Windows

1. Install Python 3.10 or newer from https://www.python.org/downloads/ (enable **Add Python to PATH** during installation).
2. Extract the Procell ZIP to a folder.
3. Open Command Prompt in that folder and run:

   ```bat
   python -m pip install -r requirements.txt
   ```

4. Double-click `run_procell.bat`, or run:

   ```bat
   python -m streamlit run app.py
   ```

5. Your browser opens Procell. Upload the mapping workbook, select a worksheet, click **Run Procell checks**, and download the consolidated report.

## Run on macOS / Linux

From the extracted folder, run:

```bash
python3 -m pip install -r requirements.txt
python3 -m streamlit run app.py
```

## Input workbook requirements

- Supported workbook types: `.xlsx` and `.xlsm`.
- The selected worksheet's column headers must be in row 1.
- Required headers (case and punctuation tolerant): `Source Table`, `Target Table`, `Target Column`, `Pseudo Logic`, `Transpose Rule`.
- Other input columns (for example `Source`, `Source Column`, `Datatype`, `Precision`, `Additional Notes`) are preserved.
- Procell ignores fully blank worksheet rows while keeping the physical Excel row number on each retained data row.

## Report worksheets

- **Summary** — key counts and interpretations.
- **Duplicate Rows** — all rows in repeated duplicate groups, including the first occurrence, with a duplicate group ID and group size.
- **Missing Mandatory** — one row per missing mandatory target column per Source-Target pair.
- **Pair Summary** — one row per Source-Target pair, including its missing-column list and pass/fail status.
- **Checked Input** — original input rows with duplicate and pair-check flags.
- **Methodology** — the rules and comparison behavior used by Procell.

## Matching rules

- Column headers and target-column values are matched case-insensitively; leading/trailing spaces are ignored.
- Duplicate keys use all four configured fields. Text comparison collapses repeated internal whitespace for matching. Original cell contents are preserved in output.
- Every row in a duplicate group of two or more rows is identified as a duplicate, including its first occurrence.
- Rows with a blank `Source Table` or `Target Table` are excluded from mandatory coverage checking, but are included in duplicate checking.

## Files

- `app.py` — Streamlit interface.
- `procell_core.py` — analysis and Excel report generation.
- `requirements.txt` — required Python packages.
- `run_procell.bat` — Windows launcher.
