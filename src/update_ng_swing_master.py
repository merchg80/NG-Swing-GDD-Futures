import argparse
import os
from datetime import datetime
from typing import Optional

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


PRODUCT_FILTER = "NG Swing GDD Futures"

# Daily file expected columns by position:
# B = HUB
# C = PRODUCT
# D = STRIP / DATE
# H = SETTLEMENT PRICE
HUB_COL_INDEX = 1          # zero-based index for column B
PRODUCT_COL_INDEX = 2      # zero-based index for column C
STRIP_COL_INDEX = 3        # zero-based index for column D
PRICE_COL_INDEX = 7        # zero-based index for column H


def normalize_date(value) -> Optional[pd.Timestamp]:
    """
    Converts Excel dates, Python dates, and date-like strings into pandas Timestamp.
    Returns None if the value cannot be parsed.
    """
    if pd.isna(value):
        return None

    parsed = pd.to_datetime(value, errors="coerce")

    if pd.isna(parsed):
        return None

    return pd.Timestamp(parsed).normalize()


def format_date_for_header(value) -> str:
    """
    Converts a parsed date to YYYY-MM-DD for the master file column header.
    """
    parsed = normalize_date(value)

    if parsed is None:
        return str(value).strip()

    return parsed.strftime("%Y-%m-%d")


def read_daily_file(daily_file: str) -> pd.DataFrame:
    """
    Reads the daily ICE download.
    This intentionally reads by column position because the file format is known:
    B = hub, C = product, D = strip/date, H = settlement price.
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
        "Strip": raw.iloc[:, STRIP_COL_INDEX],
        "SettlementPrice": raw.iloc[:, PRICE_COL_INDEX],
    })

    df["Hub"] = df["Hub"].astype(str).str.strip()
    df["Product"] = df["Product"].astype(str).str.strip()
    df["StripDate"] = df["Strip"].apply(normalize_date)
    df["DateHeader"] = df["StripDate"].apply(
        lambda x: x.strftime("%Y-%m-%d") if pd.notna(x) else None
    )

    df["SettlementPrice"] = pd.to_numeric(df["SettlementPrice"], errors="coerce")

    return df


def filter_daily_prices(
    df: pd.DataFrame,
    start_date: str,
    end_date: str,
) -> pd.DataFrame:
    """
    Filters to:
    - Product contains NG Swing GDD Futures
    - Strip date is within the requested date range
    - Hub and price are valid
    """

    start = pd.Timestamp(start_date).normalize()
    end = pd.Timestamp(end_date).normalize()

    filtered = df[
        df["Product"].str.contains(PRODUCT_FILTER, case=False, na=False)
        & df["StripDate"].notna()
        & (df["StripDate"] >= start)
        & (df["StripDate"] <= end)
        & df["Hub"].notna()
        & df["SettlementPrice"].notna()
    ].copy()

    if filtered.empty:
        raise ValueError(
            f"No rows found where Product contains '{PRODUCT_FILTER}' "
            f"between {start_date} and {end_date}."
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

    if os.path.exists(master_file):
        wb = load_workbook(master_file)
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
    Reads existing date headers from row 1.
    Returns a dictionary:
    {
        "2026-06-25": 2,
        "2026-06-26": 3,
        etc.
    }
    """

    headers = {}

    for col in range(2, ws.max_column + 1):
        value = ws.cell(row=1, column=col).value
        if value is not None and str(value).strip() != "":
            headers[str(value).strip()] = col

    return headers


def get_existing_hubs(ws):
    """
    Reads existing hubs from column A.
    Returns a dictionary:
    {
        "HENRY": 2,
        "TETCO M3": 3,
        etc.
    }
    """

    hubs = {}

    for row in range(2, ws.max_row + 1):
        value = ws.cell(row=row, column=1).value
        if value is not None and str(value).strip() != "":
            hubs[str(value).strip()] = row

    return hubs


def ensure_date_columns(ws, date_headers):
    """
    Adds any missing date headers to row 1.
    Dates are placed after existing date columns.
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
    Adds any missing hubs to column A.
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
    Writes prices into the master matrix:
    - Hub down column A
    - Dates across row 1
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
    This rebuilds the visible matrix while preserving the values.
    """

    # Read current matrix
    hubs = []
    dates = []

    for row in range(2, ws.max_row + 1):
        hub = ws.cell(row=row, column=1).value
        if hub is not None and str(hub).strip() != "":
            hubs.append(str(hub).strip())

    for col in range(2, ws.max_column + 1):
        date = ws.cell(row=1, column=col).value
        if date is not None and str(date).strip() != "":
            dates.append(str(date).strip())

    hubs = sorted(set(hubs))
    dates = sorted(set(dates))

    data = {}

    for row in range(2, ws.max_row + 1):
        hub = ws.cell(row=row, column=1).value
        if hub is None:
            continue
        hub = str(hub).strip()

        for col in range(2, ws.max_column + 1):
            date = ws.cell(row=1, column=col).value
            if date is None:
                continue
            date = str(date).strip()

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
    Applies basic formatting to the master sheet.
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
            ws.column_dimensions[col_letter].width = 22
        else:
            ws.column_dimensions[col_letter].width = 13

    ws.row_dimensions[1].height = 22

    auto_filter_range = f"A1:{get_column_letter(max_col)}{max_row}"
    ws.auto_filter.ref = auto_filter_range


def save_master(wb, output_file: str):
    """
    Saves the updated master workbook.
    """

    output_dir = os.path.dirname(output_file)

    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    wb.save(output_file)


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

    return parser.parse_args()


def main():
    args = parse_args()

    start_date = args.start
    end_date = args.end

    if not start_date:
        start_date = input("Enter start date to pull, YYYY-MM-DD: ").strip()

    if not end_date:
        end_date = input("Enter end date to pull, YYYY-MM-DD: ").strip()

    # Validate dates early
    try:
        pd.Timestamp(start_date)
        pd.Timestamp(end_date)
    except Exception as exc:
        raise ValueError("Start and end dates must be valid YYYY-MM-DD dates.") from exc

    print(f"Reading daily file: {args.daily}")
    daily_df = read_daily_file(args.daily)

    print(f"Filtering product containing: {PRODUCT_FILTER}")
    print(f"Date range: {start_date} through {end_date}")
    prices = filter_daily_prices(daily_df, start_date, end_date)

    print(f"Rows extracted: {len(prices)}")
    print(f"Hubs found: {prices['Hub'].nunique()}")
    print(f"Dates found: {prices['DateHeader'].nunique()}")

    wb, ws = create_or_load_master(args.master)

    update_master(ws, prices)
    sort_master(ws)
    format_master(ws)

    save_master(wb, args.output)

    print(f"Master file updated successfully: {args.output}")


if __name__ == "__main__":
    main()
