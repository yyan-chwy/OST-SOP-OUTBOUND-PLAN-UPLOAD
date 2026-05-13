# OST Outbound Plan Upload

This app is a Streamlit-based uploader for OB 6-6 planning data into Snowflake.

## What it does

- Connects to Snowflake using external browser authentication.
- Shows the latest upload timestamp from the target table.
- Accepts one Excel file and reads the `6to6` sheet.
- Normalizes column names and converts date fields.
- Adds an `UPLOAD_DATETIME` stamp at upload time.
- Replaces current-week rows in the destination table and inserts the new file data.
- Displays a pivot-style summary of the latest uploaded data.

## Source Code

- Main script: `OST OUTBOUND PLAN UPLOAD.py`

## Target Snowflake Table

- `EDLDB.SC_ORDER_ROUTING_TEAM_SANDBOX.OST_SOP_OUTBOUND_PLAN_UPLOAD`

## Required Excel Sheet

- Sheet name must be: `6to6`

## Required Columns in `6to6`

- `FORECAST_DATE`
- `FORECAST_WEEK_BEG`
- `PUBLISH_WEEK`
- `FULFILLMENT_CENTER_NAME`
- `ALLOCATED_UNITS_SIX_TO_SIX_UNBUFFERED`
- `BUFFER_AMOUNT`
- `SOP_PLANNED_UNITS`

The app adds this column automatically:

- `UPLOAD_DATETIME`

## Upload Logic

1. Parse `6to6` into a DataFrame.
2. Uppercase and trim all input column names.
3. Convert date-like columns to date values.
4. Add current timestamp to `UPLOAD_DATETIME`.
5. Validate required columns.
6. Delete current-week rows in Snowflake:
   - `WHERE UPLOAD_DATETIME >= DATE_TRUNC('WEEK', CURRENT_TIMESTAMP)`
7. Bulk insert new rows with `executemany`.
8. Commit transaction if successful; rollback on error.

## How to Run

1. Install dependencies:

```bash
pip install streamlit pandas snowflake-connector-python openpyxl
```

2. Start the app:

```bash
streamlit run "OST OUTBOUND PLAN UPLOAD.py"
```

3. Upload your `.xlsx` file in the UI.

## Notes

- Connection settings (user, role, schema, warehouse) are currently hardcoded in the script.
- Consider moving credentials and environment-specific settings to environment variables or Streamlit secrets.
- The delete step is week-based, so uploads overwrite rows for the current week window.