import argparse
import os
from pathlib import Path
from typing import Optional

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


PRODUCT_FILTER = "NG Swing GDD Futures"

# Daily ICE file expected columns by position:
# B = HUB
# C = PRODUCT
# D = STRIP / DATE
# H = SETTLEMENT PRICE
HUB_COL_INDEX = 1          # Column B
PRODUCT_COL_INDEX = 2      # Column C
STRIP_COL_INDEX = 3        # Column D
PRICE_COL_INDEX = 7        # Column H


def normalize_date(value) -> Optional[pd.Timestamp]:
    """
    Converts Excel dates, Python dates, and date-like strings into a normalized pandas Timestamp.
    Returns None if the value cannot be parsed.
    """
    if pd.isna(value):
        return None

    parsed = pd.to_datetime(value, errors="coerce")

    if pd.isna(parsed):
        return None

    return pd.Timestamp(parsed).normalize()


def normalize_header_date(value) -> Optional[str]:
    """
    Normalizes a master-file date header into YYYY-MM-DD.
    Handles true Excel dates, pandas timestamps, and date-like strings.
    """
    parsed = normalize_date(value)

    if parsed is None:
        if value is None:
            return None
        value_text = str(value).strip()
        return value_text if value_text else None

    return parsed.strftime("%Y-%m-%d")


def read_daily_file(daily_file: str) -> pd.DataFrame:
    """
    Reads the daily ICE download by fixed column positions:
    B = Hub
    C = Product
    D = Strip / Date
    H = Settlement Price
    """

    if not os.path.exists(daily_file):
        raise FileNotFoundError(f"Daily file not found: {daily_file}")

    raw = pd.read_excel(daily_file, header=0)

    if raw.shape[1] <= PRICE_COL_INDEX:
        raise ValueError(
            f"Daily file does not have enough columns. "
            f"Expected at least 8 columns through Column H. Found {raw.shape[1]} columns."
        )

    df = pd.DataFrame({
        "Hub": raw.iloc[:, HUB_COL_INDEX],
        "Product": raw.iloc[:, PRODUCT_COL_INDEX],
        "StripRaw": raw.iloc[:, STRIP_COL_INDEX],
        "SettlementPrice": raw.iloc[:, PRICE_COL_INDEX],
    })

    df["Hub"] = df["Hub"].astype(str).str.strip()
    df["Product"] = df["Product"].astype(str).str.strip()
    df["StripDate"] = df["StripRaw"].apply(normalize_date)

    df["DateHeader"] = df["StripDate"].apply(
        lambda x: x.strftime("%Y-%m-%d") if pd.notna(x) else None
    )

    df["SettlementPrice"] = pd.to_numeric(df["SettlementPrice"], errors="coerce")

    return df


def print_diagnostics(df: pd.DataFrame):
    """
    Prints useful diagnostics into the GitHub Actions log.
    """

    product_matches = df[
        df["Product"].str.contains(PRODUCT_FILTER, case=False, na=False)
    ].copy()

    print("")
    print("========== DIAGNOSTICS ==========")
    print(f"Total rows in file: {len(df)}")
    print(f"Rows where Product contains '{PRODUCT_FILTER}': {len(product_matches)}")

    if product_matches.empty:
        print("No matching product rows found.")
        print("Sample Product values from Column C:")
        print(df["Product"].dropna().drop_duplicates().head(25).to_string(index=False))
        print("=================================")
        return

    available_dates = sorted(
        d for d in product_matches["DateHeader"].dropna().unique()
    )

    print("")
    print("Matching Column D / Strip dates found:")
    for d in available_dates:
        print(f" - {d}")

    print("")
    print("Sample matching rows:")
    sample_cols = ["Hub", "Product", "StripRaw", "DateHeader", "SettlementPrice"]
    print(product_matches[sample_cols].head(25).to_string(index=False))
    print("=================================")
    print("")


def filter_daily_prices(
    df: pd.DataFrame,
    start_date: str,
    end_date: str,
) -> pd.DataFrame:
    """
    Filters to:
    - Product contains NG Swing GDD Futures
    - Column D / Strip date is within the requested date range
    - Hub is valid
    - Settlement price is valid
    """

    start = pd.Timestamp(start_date).normalize()
    end = pd.Timestamp(end_date).normalize()

    product_matches = df[
        df["Product"].str.contains(PRODUCT_FILTER, case=False, na=False)
    ].copy()

    if product_matches.empty:
        raise ValueError(
            f"No rows found where Column C / Product contains '{PRODUCT_FILTER}'."
        )

    available_dates = sorted(
        d for d in product_matches["DateHeader"].dropna().unique()
    )

    filtered = product_matches[
        product_matches["StripDate"].notna()
        & (product_matches["StripDate"] >= start)
        & (product_matches["StripDate"] <= end)
        & product_matches["Hub"].notna()
        & product_matches["SettlementPrice"].notna()
    ].copy()

    if filtered.empty:
        raise ValueError(
            f"No rows found for Product containing '{PRODUCT_FILTER}' "
            f"between {start_date} and {end_date} using Column D / Strip date.\n\n"
            f"Available matching Column D dates in this file are:\n"
            f"{available_dates}"
        )

    filtered = filtered[["Hub", "DateHeader", "SettlementPrice"]]

    # If duplicate hub/date records exist, keep the last one from the daily file.
    filtered = filtered.drop_duplicates(subset=["Hub", "DateHeader"], keep="last")

    return filtered


def create_or_load_master(master_file: str):
    """
    Creates a blank master workbook if one does not exist.
    Otherwise loads the existing master workbook.
    """

    master_path = Path(master_file)
    master_path.parent.mkdir(parents=True, exist_ok=True)

    if master_path.exists():
        wb = load_workbook(master_path)
        ws = wb.active
        ws.title = "NG Swing GDD"
    else:
        from openpyxl import Workbook

        wb = Workbook()
        ws = wb.active
        ws.title = "NG Swing GDD"
        ws.cell(row=1, column=1).value = "Hub"

    return wb, ws


def get_existing_headers(ws):
    """
    Reads existing date headers from Row 1 and normalizes them to YYYY-MM-DD.
    This prevents Excel date headers from being missed because they appear as datetime objects.
    """

    headers = {}

    for col in range(2, ws.max_column + 1):
        raw_value = ws.cell(row=1, column=col).value
        normalized = normalize_header_date(raw_value)

        if normalized:
            headers[normalized] = col
            ws.cell(row=1, column=col).value = normalized

    return headers


def get_existing_hubs(ws):
    """
    Reads existing hubs from Column A.
    """

    hubs = {}

    for row in range(2, ws.max_row + 1):
        value = ws.cell(row=row, column=1).value

        if value is not None and str(value).strip() != "":
            hubs[str(value).strip()] = row

    return hubs


def ensure_date_columns(ws, date_headers):
    """
    Adds missing date headers across Row 1.
    """

    existing_headers = get_existing_headers(ws)
    next_col = ws.max_column + 1 if ws.max_column >= 1 else 2

    for date_header in sorted(date_headers):
        if date_header not in existing_headers:
            ws.cell(row=1, column=next_col).value = date_header
            existing_headers[date_header] = next_col
            next_col += 1

    return existing_headers


def ensure_hub_rows(ws, hub_names):
    """
    Adds missing hubs down Column A.
    """

    existing_hubs = get_existing_hubs(ws)
    next_row = ws.max_row + 1 if ws.max_row >= 1 else 2

    if ws.cell(row=1, column=1).value is None:
        ws.cell(row=1, column=1).value = "Hub"

    for hub in sorted(hub_names):
        if hub not in existing_hubs:
            ws.cell(row=next_row, column=1).value = hub
            existing_hubs[hub] = next_row
            next_row += 1

    return existing_hubs


def update_master(ws, prices: pd.DataFrame):
    """
    Writes prices into the matrix:
    - Hubs down Column A
    - Dates across Row 1
    - Settlement prices in the body
    """

    date_headers = sorted(prices["DateHeader"].dropna().unique())
    hub_names = sorted(prices["Hub"].dropna().unique())

    header_map = ensure_date_columns(ws, date_headers)
    hub_map = ensure_hub_rows(ws, hub_names)

    for _, row in prices.iterrows():
        hub = row["Hub"]
        date_header = row["DateHeader"]
        price = row["SettlementPrice"]

        target_row = hub_map[hub]
        target_col = header_map[date_header]

        ws.cell(row=target_row, column=target_col).value = float(price)


def sort_master(ws):
    """
    Sorts dates left-to-right and hubs top-to-bottom.
    """

    hubs = []
    dates = []

    for row in range(2, ws.max_row + 1):
        hub = ws.cell(row=row, column=1).value
        if hub is not None and str(hub).strip() != "":
            hubs.append(str(hub).strip())

    for col in range(2, ws.max_column + 1):
        raw_date = ws.cell(row=1, column=col).value
        normalized_date = normalize_header_date(raw_date)
        if normalized_date:
            dates.append(normalized_date)

    hubs = sorted(set(hubs))
    dates = sorted(set(dates))

    data = {}

    for row in range(2, ws.max_row + 1):
        hub = ws.cell(row=row, column=1).value
        if hub is None:
            continue

        hub = str(hub).strip()

        for col in range(2, ws.max_column + 1):
            raw_date = ws.cell(row=1, column=col).value
            date = normalize_header_date(raw_date)

            if not date:
                continue

            value = ws.cell(row=row, column=col).value

            if value is not None:
                data[(hub, date)] = value

    # Clear sheet
    for row in ws.iter_rows():
        for cell in row:
            cell.value = None

    # Rebuild
    ws.cell(row=1, column=1).value = "Hub"

    for col_idx, date in enumerate(dates, start=2):
        ws.cell(row=1, column=col_idx).value = date

    for row_idx, hub in enumerate(hubs, start=2):
        ws.cell(row=row_idx, column=1).value = hub

        for col_idx, date in enumerate(dates, start=2):
            ws.cell(row=row_idx, column=col_idx).value = data.get((hub, date))


def format_master(ws):
    """
    Applies basic formatting.
    """

    header_fill = PatternFill("solid", fgColor="D9EAF7")
    header_font = Font(bold=True)
    thin_gray = Side(style="thin", color="D9D9D9")
    border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)

    max_row = ws.max_row
    max_col = ws.max_column

    ws.freeze_panes = "B2"

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for row in range(2, max_row + 1):
        hub_cell = ws.cell(row=row, column=1)
        hub_cell.font = Font(bold=True)
        hub_cell.fill = header_fill
        hub_cell.alignment = Alignment(horizontal="left", vertical="center")
        hub_cell.border = border

    for row in range(2, max_row + 1):
        for col in range(2, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.number_format = "0.000"
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.border = border

    for col in range(1, max_col + 1):
        col_letter = get_column_letter(col)

        if col == 1:
            ws.column_dimensions[col_letter].width = 24
        else:
            ws.column_dimensions[col_letter].width = 13

    ws.row_dimensions[1].height = 22

    if max_row >= 1 and max_col >= 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"


def save_audit_files(prices: pd.DataFrame, output_dir: str):
    """
    Saves an audit CSV so you can see exactly what was extracted.
    """

    audit_path = Path(output_dir)
    audit_path.mkdir(parents=True, exist_ok=True)

    extracted_file = audit_path / "last_extracted_ng_swing_gdd.csv"
    prices.to_csv(extracted_file, index=False)

    print(f"Audit extract saved to: {extracted_file}")


def save_master(wb, output_file: str):
    """
    Saves the updated master workbook.
    """

    output_path = Path(output_file)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def parse_args():
    parser = argparse.ArgumentParser(
        description="Update NG Swing GDD Futures master price file."
    )

    parser.add_argument(
        "--daily",
        default="input/icecleared_latest.xlsx",
        help="Path to daily ICE download file.",
    )

    parser.add_argument(
        "--master",
        default="master/master_ng_swing_gdd.xlsx",
        help="Path to existing master file. Created if it does not exist.",
    )

    parser.add_argument(
        "--output",
        default="master/master_ng_swing_gdd.xlsx",
        help="Path to save updated master file.",
    )

    parser.add_argument(
        "--start",
        default=None,
        help="Start date in YYYY-MM-DD format.",
    )

    parser.add_argument(
        "--end",
        default=None,
        help="End date in YYYY-MM-DD format.",
    )

    parser.add_argument(
        "--audit-dir",
        default="audit",
        help="Directory for audit output files.",
    )

    return parser.parse_args()


def main():
    args = parse_args()

    start_date = args.start
    end_date = args.end

    if not start_date:
        start_date = input("Enter start date to pull, YYYY-MM-DD: ").strip()

    if not end_date:
        end_date = input("Enter end date to pull, YYYY-MM-DD: ").strip()

    try:
        pd.Timestamp(start_date)
        pd.Timestamp(end_date)
    except Exception as exc:
        raise ValueError("Start and end dates must be valid YYYY-MM-DD dates.") from exc

    print(f"Reading daily file: {args.daily}")
    daily_df = read_daily_file(args.daily)

    print_diagnostics(daily_df)

    print(f"Filtering product containing: {PRODUCT_FILTER}")
    print(f"Requested Column D / Strip date range: {start_date} through {end_date}")

    prices = filter_daily_prices(daily_df, start_date, end_date)

    print("")
    print("========== EXTRACTED SUMMARY ==========")
    print(f"Rows extracted: {len(prices)}")
    print(f"Hubs found: {prices['Hub'].nunique()}")
    print(f"Dates found: {prices['DateHeader'].nunique()}")
    print("Extracted dates:")
    for d in sorted(prices["DateHeader"].unique()):
        print(f" - {d}")
    print("=======================================")
    print("")

    save_audit_files(prices, args.audit_dir)

    wb, ws = create_or_load_master(args.master)

    update_master(ws, prices)
    sort_master(ws)
    format_master(ws)

    save_master(wb, args.output)

    print(f"Master file updated successfully: {args.output}")


if __name__ == "__main__":
    main()
