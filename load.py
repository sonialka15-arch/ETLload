load data from Soruce and drop the maping after load.

"""
EXCEL TO EXCEL ETL DATA LOADER
-----------------------------
Features:
1. Read Excel source data
2. Map source columns to target columns
3. Validate mandatory columns
4. Handle missing values
5. Remove duplicate records
6. Append or replace target data
7. Generate rejected records report
8. Generate execution log and summary


"""
EXCEL TO EXCEL ETL DATA LOADER
-----------------------------
Features:
1. Read Excel source data
2. Map source columns to target columns
3. Validate mandatory columns
4. Handle missing values
5. Remove duplicate records
6. Append or replace target data
7. Generate rejected records report
8. Generate execution log and summary

Install:
pip install pandas openpyxl
"""

import pandas as pd
import logging
from pathlib import Path
from datetime import datetime


# ==========================================================
# 1. CONFIGURATION
# ==========================================================

SOURCE_FILE = "source.xlsx"
TARGET_FILE = "target.xlsx"

SOURCE_SHEET = "Sheet1"
TARGET_SHEET = "Sheet1"

REJECTED_FILE = "rejected_records.xlsx"
LOG_FILE = "excel_etl.log"

# Options: APPEND or REPLACE
LOAD_MODE = "APPEND"

# Source-to-target column mapping
COLUMN_MAPPING = {
    "Employee_ID": "Employee_ID",
    "Employee_Name": "Employee_Name",
    "Department": "Department",
    "Salary": "Salary",
    "Joining_Date": "Joining_Date"
}

# Mandatory target columns
MANDATORY_COLUMNS = [
    "Employee_ID",
    "Employee_Name"
]

# Duplicate identification columns
PRIMARY_KEY = ["Employee_ID"]


# ==========================================================
# 2. LOGGING CONFIGURATION
# ==========================================================

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)


# ==========================================================
# 3. READ SOURCE EXCEL
# ==========================================================

def read_source():

    logging.info("Reading source Excel file")

    if not Path(SOURCE_FILE).exists():
        raise FileNotFoundError(
            f"Source file not found: {SOURCE_FILE}"
        )

    df = pd.read_excel(
        SOURCE_FILE,
        sheet_name=SOURCE_SHEET
    )

    print(f"Source records: {len(df)}")

    return df


# ==========================================================
# 4. COLUMN MAPPING
# ==========================================================

def map_columns(df):

    missing = [
        col for col in COLUMN_MAPPING
        if col not in df.columns
    ]

    if missing:
        raise ValueError(
            f"Missing source columns: {missing}"
        )

    df = df[list(COLUMN_MAPPING.keys())].copy()

    df.rename(
        columns=COLUMN_MAPPING,
        inplace=True
    )

    logging.info("Column mapping completed")

    return df


# ==========================================================
# 5. DATA CLEANING
# ==========================================================

def clean_data(df):

    # Remove leading/trailing spaces from text
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].apply(
            lambda x: x.strip() if isinstance(x, str) else x
        )

    # Convert empty strings to NULL
    df = df.replace(r"^\s*$", pd.NA, regex=True)

    # Convert Salary to numeric
    if "Salary" in df.columns:
        df["Salary"] = pd.to_numeric(
            df["Salary"],
            errors="coerce"
        )

    # Convert Joining Date
    if "Joining_Date" in df.columns:
        df["Joining_Date"] = pd.to_datetime(
            df["Joining_Date"],
            errors="coerce"
        )

    logging.info("Data cleaning completed")

    return df


# ==========================================================
# 6. DATA VALIDATION
# ==========================================================

def validate_data(df):

    # Check mandatory columns
    invalid_mask = df[MANDATORY_COLUMNS].isna().any(axis=1)

    rejected = df[invalid_mask].copy()
    valid = df[~invalid_mask].copy()

    # Add rejection reason
    if not rejected.empty:
        rejected["Reject_Reason"] = (
            "Mandatory column is NULL"
        )

    logging.info(
        f"Valid records: {len(valid)}"
    )

    logging.info(
        f"Rejected records: {len(rejected)}"
    )

    return valid, rejected


# ==========================================================
# 7. REMOVE DUPLICATES
# ==========================================================

def remove_duplicates(df):

    before_count = len(df)

    df = df.drop_duplicates(
        subset=PRIMARY_KEY,
        keep="last"
    )

    removed = before_count - len(df)

    logging.info(
        f"Duplicate records removed: {removed}"
    )

    return df


# ==========================================================
# 8. LOAD TARGET EXCEL
# ==========================================================

def load_target(df):

    target_path = Path(TARGET_FILE)

    if LOAD_MODE.upper() == "APPEND" and target_path.exists():

        existing = pd.read_excel(
            TARGET_FILE,
            sheet_name=TARGET_SHEET
        )

        # Avoid inserting records already present
        if all(col in existing.columns for col in PRIMARY_KEY):

            existing_keys = existing[PRIMARY_KEY].dropna().drop_duplicates()

            source_keys = df[PRIMARY_KEY]

            new_records = df.merge(
                existing_keys,
                on=PRIMARY_KEY,
                how="left",
                indicator=True
            )

            new_records = new_records[
                new_records["_merge"] == "left_only"
            ].drop(columns=["_merge"])

            df = new_records

            logging.info(
                f"Existing target records: {len(existing)}"
            )

        final_df = pd.concat(
            [existing, df],
            ignore_index=True
        )

    elif LOAD_MODE.upper() == "REPLACE":

        final_df = df.copy()

    else:

        final_df = df.copy()

    # Write target Excel
    with pd.ExcelWriter(
        TARGET_FILE,
        engine="openpyxl",
        mode="w"
    ) as writer:

        final_df.to_excel(
            writer,
            sheet_name=TARGET_SHEET,
            index=False
        )

    logging.info(
        f"Target records after load: {len(final_df)}"
    )

    return len(df), len(final_df)


# ==========================================================
# 9. REJECTED RECORD REPORT
# ==========================================================

def generate_rejected_report(rejected):

    if not rejected.empty:

        rejected.to_excel(
            REJECTED_FILE,
            index=False
        )

        print(
            f"Rejected records saved: {REJECTED_FILE}"
        )


# ==========================================================
# 10. MAIN ETL EXECUTION
# ==========================================================

def main():

    start_time = datetime.now()

    print("\n========== EXCEL ETL STARTED ==========\n")

    try:

        # Step 1: Extract
        source_df = read_source()

        source_count = len(source_df)

        # Step 2: Transform
        mapped_df = map_columns(source_df)

        cleaned_df = clean_data(mapped_df)

        valid_df, rejected_df = validate_data(cleaned_df)

        valid_df = remove_duplicates(valid_df)

        # Step 3: Load
        loaded_count, target_count = load_target(valid_df)

        # Step 4: Rejected records
        generate_rejected_report(rejected_df)

        end_time = datetime.now()

        # Step 5: Execution summary
        print("\n========== ETL SUMMARY ==========")

        print(f"Source Records    : {source_count}")
        print(f"Valid Records     : {len(valid_df)}")
        print(f"Rejected Records  : {len(rejected_df)}")
        print(f"Loaded Records    : {loaded_count}")
        print(f"Target Total      : {target_count}")
        print(f"Load Mode         : {LOAD_MODE}")
        print(f"Start Time        : {start_time}")
        print(f"End Time          : {end_time}")
        print(f"Duration          : {end_time - start_time}")

        print("\nETL PROCESS COMPLETED SUCCESSFULLY")

        logging.info("ETL process completed successfully")

    except Exception as e:

        logging.exception("ETL PROCESS FAILED")

        print("\nETL PROCESS FAILED")
        print(f"Error: {e}")


# ==========================================================
# 11. EXECUTION ENTRY POINT
# ==========================================================

if __name__ == "__main__":
    main()
