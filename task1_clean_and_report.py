#!/usr/bin/env python3
"""
Task 1: Data Handling (Lead Management Basics)
- Reads a leads CSV
- Validates emails
- Removes duplicates
- Writes clean_customers.csv
- Generates leads_report.xlsx with daily and weekly lead counts and a summary

Usage:
    python task1_clean_and_report.py --input leads.csv --out clean_customers.csv --report leads_report.xlsx

Expected minimum columns in leads CSV:
    - name (optional)
    - email
    - phone (optional)
    - date (date when lead was captured) -- if missing, script will try to infer or mark as NaT
"""

import argparse
import re
import sys
from pathlib import Path

import pandas as pd

EMAIL_REGEX = re.compile(
    # simpler robust regex for common valid emails (not perfect but practical)
    r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)"
)


def valid_email(e):
    if pd.isna(e):
        return False
    return EMAIL_REGEX.match(str(e).strip()) is not None


def parse_args():
    p = argparse.ArgumentParser(description="Clean leads and generate reports")
    p.add_argument("--input", "-i", required=True, help="Input leads CSV file")
    p.add_argument("--out", "-o", default="clean_customers.csv", help="Output clean CSV")
    p.add_argument("--report", "-r", default="leads_report.xlsx", help="Output Excel report")
    return p.parse_args()


def load_csv(path):
    try:
        df = pd.read_csv(path, dtype=str)  # read everything as str to avoid surprises
        return df
    except Exception as e:
        print(f"Error reading CSV {path}: {e}", file=sys.stderr)
        raise


def clean_leads(df):
    # normalize column names
    df = df.rename(columns={c: c.strip() for c in df.columns})
    # ensure email column exists
    email_col = None
    for c in df.columns:
        if c.lower() == "email":
            email_col = c
            break
    if email_col is None:
        raise ValueError("Input CSV must contain an 'email' column")

    # Trim whitespace
    df[email_col] = df[email_col].astype(str).str.strip()

    # Validate emails
    df["__valid_email"] = df[email_col].apply(valid_email)
    valid_df = df[df["__valid_email"]].copy()
    invalid_count = len(df) - len(valid_df)

    # Drop duplicates based on email (keep first)
    before = len(valid_df)
    valid_df = valid_df.drop_duplicates(subset=[email_col])
    removed_dups = before - len(valid_df)

    # Parse date column if present
    date_col = None
    for c in df.columns:
        if c.lower() in ("date", "created_at", "lead_date"):
            date_col = c
            break
    if date_col:
        valid_df[date_col] = pd.to_datetime(valid_df[date_col], errors="coerce")
    else:
        # no date column — create one with NaT
        valid_df["date"] = pd.NaT
        date_col = "date"

    # Add helper columns: day, week
    valid_df["lead_day"] = valid_df[date_col].dt.date
    valid_df["lead_week"] = valid_df[date_col].dt.to_period("W").apply(lambda r: r.start_time.date() if pd.notna(r) else pd.NaT)

    # cleanup marker
    valid_df = valid_df.drop(columns=["__valid_email"])
    return valid_df, invalid_count, removed_dups


def generate_report(df, email_col, report_path):
    # daily counts (count by date)
    # if date not available for some rows, they will be excluded from day/week counts but included in unique counts
    if "lead_day" in df.columns:
        daily = df.groupby("lead_day")[email_col].count().rename("count").reset_index()
    else:
        daily = pd.DataFrame(columns=["lead_day", "count"])

    # weekly counts
    if "lead_week" in df.columns:
        weekly = df.groupby("lead_week")[email_col].count().rename("count").reset_index()
    else:
        weekly = pd.DataFrame(columns=["lead_week", "count"])

    unique_customers = df[email_col].nunique()
    total_leads = len(df)

    # Write to Excel
    with pd.ExcelWriter(report_path, engine="xlsxwriter") as writer:
        daily.to_excel(writer, sheet_name="Daily Leads", index=False)
        weekly.to_excel(writer, sheet_name="Weekly Leads", index=False)
        summary = pd.DataFrame(
            [
                {"Metric": "Total Clean Leads", "Value": total_leads},
                {"Metric": "Unique Customers (by email)", "Value": unique_customers},
            ]
        )
        summary.to_excel(writer, sheet_name="Summary", index=False)

    return {"daily_rows": len(daily), "weekly_rows": len(weekly), "unique_customers": unique_customers}


def main():
    args = parse_args()
    in_path = Path(args.input)
    out_path = Path(args.out)
    report_path = Path(args.report)

    df = load_csv(in_path)
    try:
        cleaned, invalid_count, removed_dups = clean_leads(df)
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(2)

    # find canonical email column name again
    email_col = None
    for c in cleaned.columns:
        if c.lower() == "email":
            email_col = c
            break

    cleaned.to_csv(out_path, index=False)
    meta = generate_report(cleaned, email_col, report_path)

    print("=== Task 1: Completed ===")
    print(f"Input file: {in_path}")
    print(f"Clean file written: {out_path} ({len(cleaned)} rows)")
    print(f"Invalid emails removed: {invalid_count}")
    print(f"Duplicate emails removed: {removed_dups}")
    print(f"Excel report written: {report_path}")
    print(f"Unique customers (by email): {meta['unique_customers']}")


if __name__ == "__main__":
    main()
