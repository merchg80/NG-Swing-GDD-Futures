from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(
    page_title="ICE Settlement History",
    layout="wide",
)

ROOT = Path(__file__).resolve().parents[1]
INPUT = ROOT / "input"

IDENTITY = [
    "contract",
    "product",
    "hub",
    "strip",
    "contract_type",
    "strike",
    "expiration_date",
    "product_id",
]

REQUIRED = [
    "trade_date",
    "hub",
    "product",
    "strip",
    "contract",
    "contract_type",
    "strike",
    "settlement_price",
    "net_change",
    "expiration_date",
    "product_id",
]


@st.cache_data(show_spinner=False, max_entries=1)
def read_file(filename, modified, size):
    """Cache one workbook; file metadata invalidates changed files."""
    df = pd.read_excel(
        filename,
        engine="openpyxl",
        dtype=object,
    )

    df.columns = [
        str(column).strip().lower().replace(" ", "_")
        for column in df.columns
    ]

    missing = set(REQUIRED) - set(df.columns)
    if missing:
        raise ValueError(
            f"{Path(filename).name}: missing columns {sorted(missing)}"
        )

    df = df[REQUIRED].dropna(how="all").copy()

    for column in [
        "contract", "product", "hub", "contract_type", "product_id"
    ]:
        df[column] = (
            df[column].fillna("").astype(str).str.strip()
        )

    df["contract"] = df["contract"].str.upper()

    for column in ["trade_date", "strip", "expiration_date"]:
        original = df[column]
        parsed = pd.to_datetime(original, errors="coerce")

        if (original.notna() & parsed.isna()).any():
            raise ValueError(
                f"{Path(filename).name}: invalid date in {column}"
            )

        df[column] = parsed

    for column in ["strike", "settlement_price", "net_change"]:
        original = df[column]
        parsed = pd.to_numeric(original, errors="coerce")

        invalid = (
            original.notna()
            & original.astype(str).str.strip().ne("")
            & parsed.isna()
        )

        if invalid.any():
            raise ValueError(
                f"{Path(filename).name}: invalid number in {column}"
            )

        df[column] = parsed

    if (
        df["trade_date"].isna()
        | df["strip"].isna()
        | df["contract"].eq("")
    ).any():
        raise ValueError(
            f"{Path(filename).name}: missing trade date, strip, or code"
        )

    df["source_file"] = str(Path(filename).relative_to(ROOT))

    return df


@st.cache_data(show_spinner=False, max_entries=1)
def load_history(manifest, codes):
    frames = []

    for filename, modified, size in manifest:
        df = read_file(filename, modified, size)

        # Filter before combining the historical files.
        df = df[df["contract"].isin(codes)].copy()

        if not df.empty:
            frames.append(df)

    if not frames:
        raise ValueError("No rows matched the selected ICE codes.")

    data = pd.concat(frames, ignore_index=True)

    key = ["trade_date"] + IDENTITY

    counts = (
        data.groupby(key, dropna=False)[
            ["settlement_price", "net_change"]
        ]
        .nunique(dropna=False)
    )

    conflicting = counts[(counts > 1).any(axis=1)]

    if not conflicting.empty:
        raise ValueError(
            f"{len(conflicting)} duplicate contract/date keys have "
            "conflicting prices. Resolve the source files before "
            "calculating averages."
        )

    sources = (
        data.groupby(key, dropna=False)["source_file"]
        .agg(lambda values: "; ".join(sorted(set(values))))
        .reset_index()
    )

    data = (
        data.drop(columns="source_file")
        .drop_duplicates(subset=key)
        .merge(sources, on=key, how="left", validate="one_to_one")
        .sort_values(["contract", "strip", "trade_date"])
        .reset_index(drop=True)
    )

    return data


def choose_many(label, frame, column, default=None):
    options = sorted(frame[column].dropna().unique().tolist())

    if not options:
        return frame

    selected = st.sidebar.multiselect(
        label,
        options,
        default=default if default is not None else options,
    )

    return frame[frame[column].isin(selected)]


def summarize(frame, metric, monthly=False):
    work = frame.copy()
    keys = IDENTITY.copy()

    if monthly:
        work["trade_month"] = work["trade_date"].dt.strftime("%Y-%m")
        keys.append("trade_month")

    return (
        work.groupby(keys, dropna=False)[metric]
        .agg(
            average="mean",
            observations="count",
            minimum="min",
            maximum="max",
        )
        .reset_index()
    )


def csv_bytes(frame):
    return frame.to_csv(index=False).encode("utf-8-sig")


st.title("ICE Settlement History")

st.caption(
    "Filter historical settlements by ICE code, product, hub, "
    "delivery strip, and trading dates."
)

with st.sidebar:
    st.header("Import settings")

    code_text = st.text_input(
        "ICE codes to load",
        value="FQH,UCS,WAH,NTO,TGB,DOM,H,FQT,TPH,HXS",
        help="Comma-separated ICE contract codes.",
    )

    st.caption(
        "Loading selected codes keeps the app smaller. "
        "Add other codes here when needed."
    )

    refresh = st.button("Reload available files")

if refresh:
    st.cache_data.clear()
    st.rerun()

codes = tuple(sorted({
    item.strip().upper()
    for item in code_text.split(",")
    if item.strip()
}))

if not codes:
    st.info("Enter at least one ICE code in the sidebar.")
    st.stop()

files = sorted(
    path for path in INPUT.rglob("icecleared*.xlsx")
    if not path.name.startswith("~$")
)

if not files:
    st.error("No icecleared*.xlsx files were found in input.")
    st.stop()

manifest = tuple(
    (str(path), path.stat().st_mtime_ns, path.stat().st_size)
    for path in files
)

try:
    with st.spinner("Loading and checking settlement history..."):
        data = load_history(manifest, codes)
except Exception as error:
    st.error(str(error))
    st.stop()

st.caption(
    f"{len(files):,} files found | "
    f"{len(data):,} unique matching rows | "
    f"Trade dates: {data['trade_date'].min():%b %d, %Y} "
    f"through {data['trade_date'].max():%b %d, %Y}"
)

missing_codes = sorted(set(codes) - set(data["contract"].unique()))
if missing_codes:
    st.warning("Codes not found: " + ", ".join(missing_codes))

st.sidebar.header("Report filters")

available_codes = sorted(data["contract"].unique().tolist())
preferred = next(
    (code for code in ["HXS", "FQT", "TPH"] if code in available_codes),
    available_codes[0],
)

view = choose_many(
    "ICE code",
    data,
    "contract",
    default=[preferred],
)

view = choose_many("Product", view, "product")
view = choose_many("Hub", view, "hub")

types = sorted(view["contract_type"].dropna().unique().tolist())

view = choose_many(
    "Contract type",
    view,
    "contract_type",
    default=["F"] if "F" in types else types,
)

available_strips = sorted(view["strip"].dropna().unique().tolist())

if not available_strips:
    st.info("No delivery strips match these filters.")
    st.stop()

chosen_strips = st.sidebar.multiselect(
    "Delivery strip",
    available_strips,
    default=available_strips[:1],
    format_func=lambda value: pd.Timestamp(value).strftime("%b %Y"),
)

view = view[view["strip"].isin(chosen_strips)]

metric_label = st.sidebar.selectbox(
    "Price field",
    ["Settlement price", "Net change"],
)

metric = (
    "settlement_price"
    if metric_label == "Settlement price"
    else "net_change"
)

first_date = data["trade_date"].min().date()
last_date = data["trade_date"].max().date()

date_range = st.sidebar.date_input(
    "Trade-date range",
    value=(first_date, last_date),
    min_value=first_date,
    max_value=last_date,
)

if len(date_range) != 2:
    st.info("Select both a first and last trade date.")
    st.stop()

start, end = date_range

view = view[
    view["trade_date"].between(
        pd.Timestamp(start),
        pd.Timestamp(end),
    )
].copy()

if view.empty:
    st.info("No records match the selected filters.")
    st.stop()

period = summarize(view, metric)
monthly = summarize(view, metric, monthly=True)

left, middle, right = st.columns(3)
left.metric("Matching records", f"{len(view):,}")
middle.metric(
    "Numeric observations",
    f"{view[metric].count():,}",
)
right.metric("Separate price series", f"{len(period):,}")

st.caption(
    "Averages use available numeric observations. Missing prices "
    "are excluded; zero prices are included. Results are separated "
    "by code, product, hub, strip, type, strike, expiration, and "
    "product ID. Trading dates with no records are not filled."
)

chart_view = st.radio(
    "Chart",
    ["Daily prices", "Monthly averages"],
    horizontal=True,
)

chart_data = view.copy() if chart_view == "Daily prices" else monthly.copy()

# Generate labels containing every identity field.
chart_data["series"] = chart_data[IDENTITY].apply(
    lambda row: " | ".join(
        f"{column}={row[column]}"
        for column in IDENTITY
    ),
    axis=1,
)

series_count = chart_data["series"].nunique()

if series_count > 30:
    st.info(
        "More than 30 series selected. Narrow the filters to "
        "display the chart. All matching results remain in the tables."
    )
else:
    if chart_view == "Daily prices":
        x_column = "trade_date"
        y_column = metric
    else:
        x_column = "trade_month"
        y_column = "average"

    chart_data = chart_data.sort_values(x_column)

    fig = px.line(
        chart_data,
        x=x_column,
        y=y_column,
        color="series",
        markers=True,
        labels={
            x_column: "Trade date" if chart_view == "Daily prices"
            else "Trade month",
            y_column: metric_label,
            "series": "Contract series",
        },
    )

    fig.update_layout(
        height=550,
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.25,
        ),
    )

    st.plotly_chart(fig, use_container_width=True)

tab1, tab2, tab3 = st.tabs([
    "Selected-period averages",
    "Monthly averages",
    "Underlying prices",
])

with tab1:
    st.dataframe(period, use_container_width=True, hide_index=True)
    st.download_button(
        "Download period averages CSV",
        csv_bytes(period),
        "ice_period_averages.csv",
        "text/csv",
    )

with tab2:
    st.dataframe(monthly, use_container_width=True, hide_index=True)
    st.download_button(
        "Download monthly averages CSV",
        csv_bytes(monthly),
        "ice_monthly_averages.csv",
        "text/csv",
    )

with tab3:
    prices = view.sort_values(
        ["trade_date", "contract", "strip"]
    )
    st.dataframe(prices, use_container_width=True, hide_index=True)
    st.download_button(
        "Download underlying prices CSV",
        csv_bytes(prices),
        "ice_prices.csv",
        "text/csv",
    )
