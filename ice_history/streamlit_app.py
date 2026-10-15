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
def load_history(manifest, codes):
    """
    Load historical files.

    Empty codes means every instrument.
    File metadata invalidates the cache when files change.
    """
    frames = []

    for filename, modified, size in manifest:
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
                f"{Path(filename).name}: missing columns "
                f"{sorted(missing)}"
            )

        df = df[REQUIRED].dropna(how="all").copy()

        for column in [
            "contract",
            "product",
            "hub",
            "contract_type",
            "product_id",
        ]:
            df[column] = (
                df[column]
                .fillna("")
                .astype(str)
                .str.strip()
            )

        df["contract"] = df["contract"].str.upper()

        # Blank import field means load every ICE code.
        if codes:
            df = df[df["contract"].isin(codes)].copy()

        if df.empty:
            continue

        for column in [
            "trade_date",
            "strip",
            "expiration_date",
        ]:
            original = df[column]

            nonblank = (
                original.notna()
                & original.astype(str).str.strip().ne("")
            )

            parsed = pd.to_datetime(
                original.where(nonblank),
                errors="coerce",
            )

            invalid = nonblank & parsed.isna()

            if invalid.any():
                examples = original[invalid].head(3).tolist()

                raise ValueError(
                    f"{Path(filename).name}: invalid date "
                    f"in {column}: {examples}"
                )

            df[column] = parsed

        for column in [
            "strike",
            "settlement_price",
            "net_change",
        ]:
            original = df[column]

            nonblank = (
                original.notna()
                & original.astype(str).str.strip().ne("")
            )

            parsed = pd.to_numeric(
                original.where(nonblank),
                errors="coerce",
            )

            invalid = nonblank & parsed.isna()

            if invalid.any():
                raise ValueError(
                    f"{Path(filename).name}: invalid "
                    f"number in {column}"
                )

            df[column] = parsed

        bad_identity = (
            df["trade_date"].isna()
            | df["strip"].isna()
            | df["contract"].eq("")
        )

        if bad_identity.any():
            raise ValueError(
                f"{Path(filename).name}: "
                f"{int(bad_identity.sum())} rows are missing "
                "trade date, delivery strip, or ICE code."
            )

        df["source_file"] = str(
            Path(filename).relative_to(ROOT)
        )

        frames.append(df)

    if not frames:
        raise ValueError(
            "No pricing records matched. Leave the ICE codes "
            "field blank to load every instrument."
        )

    data = pd.concat(frames, ignore_index=True)
    del frames

    key = ["trade_date"] + IDENTITY

    duplicate_mask = data.duplicated(
        subset=key,
        keep=False,
    )

    duplicate_rows = data.loc[duplicate_mask]

    if not duplicate_rows.empty:
        counts = (
            duplicate_rows.groupby(
                key,
                dropna=False,
                observed=True,
            )[["settlement_price", "net_change"]]
            .nunique(dropna=False)
        )

        conflicting = counts[(counts > 1).any(axis=1)]

        if not conflicting.empty:
            examples = (
                conflicting.reset_index()[
                    ["trade_date", "contract", "hub", "strip"]
                ]
                .head(5)
                .to_string(index=False)
            )

            raise ValueError(
                f"{len(conflicting)} duplicate contract/date "
                "keys contain conflicting prices. Resolve the "
                "source records before calculating averages.\n\n"
                f"Examples:\n{examples}"
            )

        # Combine the filenames for identical duplicate records.
        unique_rows = data.loc[~duplicate_mask].copy()

        sources = (
            duplicate_rows.groupby(
                key,
                dropna=False,
                observed=True,
            )["source_file"]
            .agg(
                lambda values: "; ".join(
                    sorted(set(values))
                )
            )
            .reset_index()
        )

        deduplicated = (
            duplicate_rows.drop(columns="source_file")
            .drop_duplicates(subset=key)
            .merge(
                sources,
                on=key,
                how="left",
                validate="one_to_one",
            )
        )

        data = pd.concat(
            [unique_rows, deduplicated],
            ignore_index=True,
        )

    data = data.sort_values(
        ["contract", "hub", "strip", "trade_date"]
    ).reset_index(drop=True)

    # Repeated descriptions and identifiers use less memory
    # as categorical columns.
    for column in [
        "contract",
        "product",
        "hub",
        "contract_type",
        "product_id",
        "source_file",
    ]:
        data[column] = data[column].astype("category")

    return data


def choose_many(label, frame, column, default=None):
    """
    Empty selection means all available values.
    Selections cascade from the preceding filters.
    """
    options = sorted(
        frame[column].dropna().unique().tolist()
    )

    if not options:
        return frame

    selected = st.sidebar.multiselect(
        label,
        options=options,
        default=default or [],
        help="Leave empty to include all available values.",
    )

    if not selected:
        return frame

    return frame[frame[column].isin(selected)]


def summarize(frame, metric, monthly=False):
    work = frame.copy()
    keys = IDENTITY.copy()

    if monthly:
        work["trade_month"] = (
            work["trade_date"].dt.strftime("%Y-%m")
        )
        keys.append("trade_month")

    return (
        work.groupby(
            keys,
            dropna=False,
            observed=True,
        )[metric]
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
    "Access historical instruments from the settlement files. "
    "Filter by ICE code, product, hub, delivery strip, "
    "and trading dates."
)

with st.sidebar:
    st.header("Import settings")

    # Blank by default: every instrument is available.
    code_text = st.text_input(
        "ICE codes to load",
        value="",
        help=(
            "Leave blank to load every ICE code. "
            "Optionally enter comma-separated codes "
            "to reduce the import size."
        ),
    )

    st.caption(
        "Blank loads every instrument in the files. "
        "Use the report filters below to choose what "
        "you want to analyze."
    )

    refresh = st.button("Reload available files")

    st.caption(
        "Reload rereads files available on this server. "
        "It does not pull GitHub changes itself."
    )

if refresh:
    st.cache_data.clear()
    st.rerun()

codes = tuple(
    sorted(
        {
            item.strip().upper()
            for item in code_text.split(",")
            if item.strip()
        }
    )
)

files = sorted(
    path
    for path in INPUT.rglob("icecleared*.xlsx")
    if not path.name.startswith("~$")
)

if not files:
    st.error(
        "No icecleared*.xlsx files were found "
        "in the repository's input folder."
    )
    st.stop()

manifest = tuple(
    (
        str(path),
        path.stat().st_mtime_ns,
        path.stat().st_size,
    )
    for path in files
)

try:
    with st.spinner(
        "Loading and checking settlement history. "
        "The first import may take several minutes..."
    ):
        data = load_history(manifest, codes)

except Exception as error:
    st.error(str(error))
    st.stop()

st.caption(
    f"{len(files):,} source files | "
    f"{len(data):,} unique loaded records | "
    f"{data['contract'].nunique():,} ICE codes | "
    f"Trade dates: "
    f"{data['trade_date'].min():%b %d, %Y} "
    f"through {data['trade_date'].max():%b %d, %Y}"
)

if codes:
    missing_codes = sorted(
        set(codes)
        - set(data["contract"].dropna().unique())
    )

    if missing_codes:
        st.warning(
            "Codes not found: " + ", ".join(missing_codes)
        )

st.sidebar.header("Report filters")

st.sidebar.caption(
    "An empty report filter means all available values. "
    "Options narrow as you make selections."
)

available_codes = sorted(
    data["contract"].dropna().unique().tolist()
)

preferred = next(
    (
        code
        for code in ["HXS", "FQT", "TPH"]
        if code in available_codes
    ),
    available_codes[0],
)

view = choose_many(
    "ICE code",
    data,
    "contract",
    default=[preferred],
)

view = choose_many(
    "Product",
    view,
    "product",
)

view = choose_many(
    "Hub",
    view,
    "hub",
)

# No default type exclusion: futures and options remain accessible.
view = choose_many(
    "Contract type",
    view,
    "contract_type",
)

available_strips = sorted(
    view["strip"].dropna().unique().tolist()
)

if not available_strips:
    st.info("No delivery strips match these filters.")
    st.stop()

chosen_strips = st.sidebar.multiselect(
    "Delivery strip",
    options=available_strips,
    default=available_strips[:1],
    format_func=lambda value: (
        pd.Timestamp(value).strftime("%b %d, %Y")
    ),
    help=(
        "Select one or more delivery periods. "
        "Clear the selection to include all."
    ),
)

if chosen_strips:
    view = view[
        view["strip"].isin(chosen_strips)
    ]

# These additional filters allow individual option
# contracts and expirations to be isolated.
with st.sidebar.expander("Additional contract filters"):
    strikes = sorted(
        view["strike"].dropna().unique().tolist()
    )

    selected_strikes = st.multiselect(
        "Strike",
        options=strikes,
        help=(
            "Leave empty for all strikes, including "
            "records with no strike."
        ),
    )

    if selected_strikes:
        view = view[
            view["strike"].isin(selected_strikes)
        ]

    expirations = sorted(
        view["expiration_date"]
        .dropna()
        .unique()
        .tolist()
    )

    selected_expirations = st.multiselect(
        "Expiration date",
        options=expirations,
        format_func=lambda value: (
            pd.Timestamp(value).strftime("%b %d, %Y")
        ),
        help="Leave empty for all expirations.",
    )

    if selected_expirations:
        view = view[
            view["expiration_date"].isin(
                selected_expirations
            )
        ]

metric_label = st.sidebar.selectbox(
    "Price field",
    options=[
        "Settlement price",
        "Net change",
    ],
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
    st.info(
        "Select both a first and last trade date."
    )
    st.stop()

start, end = date_range

if start > end:
    st.error(
        "The first trade date must precede "
        "the last trade date."
    )
    st.stop()

view = view[
    view["trade_date"].between(
        pd.Timestamp(start),
        pd.Timestamp(end),
    )
].copy()

if view.empty:
    st.info(
        "No records match the selected filters."
    )
    st.stop()

period = summarize(
    view,
    metric,
)

monthly = summarize(
    view,
    metric,
    monthly=True,
)

left, middle, right = st.columns(3)

left.metric(
    "Matching records",
    f"{len(view):,}",
)

middle.metric(
    "Numeric observations",
    f"{view[metric].count():,}",
)

right.metric(
    "Separate price series",
    f"{len(period):,}",
)

st.caption(
    "Averages are arithmetic averages of available "
    "observations, not volume-weighted. Missing prices "
    "are excluded; zero prices are included. Each code, "
    "product, hub, strip, type, strike, expiration, "
    "and product ID is calculated separately. "
    "Missing trading dates are not filled."
)

chart_view = st.radio(
    "Chart view",
    options=[
        "Daily prices",
        "Monthly averages",
    ],
    horizontal=True,
)

if len(period) > 30:
    st.info(
        "More than 30 price series are selected. "
        "Narrow the report filters to display the chart. "
        "All matching results remain available "
        "in the tables."
    )

else:
    chart_data = (
        view.copy()
        if chart_view == "Daily prices"
        else monthly.copy()
    )

    chart_data["series"] = chart_data[
        IDENTITY
    ].apply(
        lambda row: " | ".join(
            f"{column}={row[column]}"
            for column in IDENTITY
        ),
        axis=1,
    )

    if chart_view == "Daily prices":
        x_column = "trade_date"
        y_column = metric
        x_label = "Trade date"

    else:
        x_column = "trade_month"
        y_column = "average"
        x_label = "Trade month"

    chart_data = chart_data.sort_values(
        x_column
    )

    fig = px.line(
        chart_data,
        x=x_column,
        y=y_column,
        color="series",
        markers=True,
        labels={
            x_column: x_label,
            y_column: metric_label,
            "series": "Contract series",
        },
    )

    fig.update_traces(
        connectgaps=False,
    )

    fig.update_layout(
        height=550,
        legend={
            "orientation": "h",
            "yanchor": "top",
            "y": -0.25,
        },
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
    )

tab1, tab2, tab3 = st.tabs(
    [
        "Selected-period averages",
        "Monthly averages",
        "Underlying prices",
    ]
)

with tab1:
    st.caption(
        "One average per distinct price series "
        "across the selected trading-date range."
    )

    st.dataframe(
        period,
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        "Download period averages CSV",
        data=csv_bytes(period),
        file_name="ice_period_averages.csv",
        mime="text/csv",
    )

with tab2:
    st.caption(
        "Average of daily observations within "
        "each trading month, keeping delivery "
        "strips and contract identities separate."
    )

    st.dataframe(
        monthly,
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        "Download monthly averages CSV",
        data=csv_bytes(monthly),
        file_name="ice_monthly_averages.csv",
        mime="text/csv",
    )

with tab3:
    prices = view.sort_values(
        [
            "trade_date",
            "contract",
            "strip",
        ]
    )

    st.dataframe(
        prices,
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        "Download underlying prices CSV",
        data=csv_bytes(prices),
        file_name="ice_prices.csv",
        mime="text/csv",
    )
