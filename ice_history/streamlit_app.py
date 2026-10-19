from pathlib import Path

import duckdb
import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(
    page_title="Natural Gas Pricing Analysis",
    layout="wide",
)

ROOT = Path(__file__).resolve().parents[1]
ARCHIVE = ROOT / "ice_archive"

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

FAMILY = [
    "contract",
    "product",
    "hub",
    "contract_type",
    "strike",
    "product_id",
]


def file_signature(path):
    stat = path.stat()
    return (
        str(path),
        stat.st_mtime_ns,
        stat.st_size,
    )


@st.cache_data(max_entries=2, show_spinner=False)
def read_catalog(signature):
    return pd.read_parquet(signature[0])


@st.cache_data(max_entries=2, show_spinner=False)
def read_master(signature):
    df = pd.read_csv(signature[0])
    df.columns = [str(c).strip() for c in df.columns]

    date_column = next(
        (c for c in df.columns if c.lower() == "date"),
        None,
    )

    if date_column is None:
        raise ValueError("The master CSV has no Date column.")

    df = df.rename(columns={date_column: "Date"})
    df["Date"] = pd.to_datetime(df["Date"], errors="raise")

    if df["Date"].isna().any():
        raise ValueError("The master CSV contains missing dates.")

    for column in df.columns:
        if column != "Date":
            df[column] = pd.to_numeric(
                df[column],
                errors="raise",
            )

    df = df.drop_duplicates()

    if df["Date"].duplicated().any():
        raise ValueError(
            "The master CSV has conflicting records for one date."
        )

    return df.sort_values("Date").set_index("Date")


@st.cache_data(max_entries=4, show_spinner=False)
def query_prices(
    archive_signature,
    instruments,
    trade_start,
    trade_end,
):
    connection = duckdb.connect()

    try:
        connection.execute("SET memory_limit='512MB'")
        connection.execute("SET threads=2")

        connection.read_parquet(
            str(ARCHIVE / "parts" / "*.parquet"),
            union_by_name=True,
        ).create_view("prices")

        wanted = pd.DataFrame(
            instruments,
            columns=IDENTITY,
        )

        # Explicit types also handle selections with all-null strikes
        # or expiration dates.
        for column in ["strip", "expiration_date"]:
            wanted[column] = pd.to_datetime(wanted[column])

        wanted["strike"] = pd.to_numeric(
            wanted["strike"],
            errors="raise",
        ).astype("float64")

        connection.register("wanted_instruments", wanted)

        join_conditions = " AND ".join(
            f'p."{column}" IS NOT DISTINCT FROM w."{column}"'
            for column in IDENTITY
        )

        result = connection.execute(
            f"""
            SELECT DISTINCT p.*
            FROM prices AS p
            INNER JOIN wanted_instruments AS w
              ON {join_conditions}
            WHERE p.trade_date >= CAST(? AS DATE)
              AND p.trade_date <= CAST(? AS DATE)
            """,
            [str(trade_start), str(trade_end)],
        ).df()

        return result

    finally:
        connection.close()


def date_controls(
    first,
    last,
    key,
    start_label="Start date",
    end_label="End date",
):
    first = pd.Timestamp(first).date()
    last = pd.Timestamp(last).date()

    # Keep saved dates within the current data's bounds.
    for suffix, default in [
        ("_start", first),
        ("_end", last),
    ]:
        state_key = key + suffix
        current = st.session_state.get(state_key, default)

        if not first <= current <= last:
            st.session_state[state_key] = default

    left, right = st.columns(2)

    start = left.date_input(
        start_label,
        value=first,
        min_value=first,
        max_value=last,
        key=key + "_start",
    )

    end = right.date_input(
        end_label,
        value=last,
        min_value=first,
        max_value=last,
        key=key + "_end",
    )

    if start > end:
        st.error(
            f"{start_label} must be on or before {end_label.lower()}."
        )
        st.stop()

    return start, end


def multi_filter(frame, column, label, key):
    options = sorted(
        frame[column].dropna().unique().tolist()
    )

    if not options:
        return frame

    # Remove selections that are unavailable after upstream changes.
    if key in st.session_state:
        st.session_state[key] = [
            value
            for value in st.session_state[key]
            if value in options
        ]

    selected = st.sidebar.multiselect(
        label,
        options=options,
        key=key,
        help="Leave empty to include all available values.",
    )

    if selected:
        return frame[frame[column].isin(selected)]

    return frame


def instrument_tuple(row):
    return tuple(
        None if pd.isna(value) else value
        for value in row
    )


def show_report(
    wide,
    key,
    metric_label,
    strip_average=False,
    expected_columns=None,
):
    wide = wide.sort_index().copy()

    if expected_columns:
        # Retain requested delivery columns even if all their
        # prices are missing within the chosen trade dates.
        wide = wide.reindex(columns=expected_columns)

    if wide.empty or wide.shape[1] == 0:
        st.info("No records match this selection.")
        return

    if strip_average:
        include_average = st.checkbox(
            "Show selected-strip average",
            value=True,
            key=key + "_average",
        )

        require_all = st.checkbox(
            "Require every selected delivery column for the average",
            value=True,
            key=key + "_complete",
            help=(
                "A daily strip average is blank when a selected "
                "delivery column has no numeric price on that date."
            ),
        )

        if include_average:
            average = wide.mean(axis=1)

            if require_all:
                average = average.where(
                    wide.notna().all(axis=1)
                )

            wide["Selected-strip average"] = average

    means = wide.mean()

    monthly = wide.groupby(
        wide.index.to_period("M")
    ).mean()

    frequency = st.radio(
        "Chart frequency",
        options=["Daily", "Monthly"],
        horizontal=True,
        key=key + "_frequency",
    )

    chart = wide.copy() if frequency == "Daily" else monthly.copy()

    if frequency == "Monthly":
        chart.index = chart.index.to_timestamp()

    if chart.shape[1] <= 30:
        chart.index.name = "Date"

        long = chart.reset_index().melt(
            id_vars="Date",
            var_name="Series",
            value_name="Value",
        )

        figure = px.line(
            long,
            x="Date",
            y="Value",
            color="Series",
            markers=True,
            labels={"Value": metric_label},
        )

        figure.update_traces(connectgaps=False)

        for trace in figure.data:
            if trace.name == "Selected-strip average":
                trace.line.update(
                    color="black",
                    width=4,
                    dash="dash",
                )

        figure.update_layout(
            height=520,
            legend={
                "orientation": "h",
                "y": -0.25,
                "yanchor": "top",
            },
        )

        st.plotly_chart(
            figure,
            use_container_width=True,
        )

    else:
        st.info(
            "This selection contains more than 30 chart lines. "
            "Narrow the delivery range to show the chart. "
            "All selected columns remain in the tables."
        )

    daily = wide.copy()
    daily.index = daily.index.strftime("%Y-%m-%d")
    daily.index.name = "Date"
    daily.loc["AVERAGE OVER SELECTED DATES"] = means

    monthly.index = monthly.index.astype(str)
    monthly.index.name = "Month"
    monthly.loc["AVERAGE OVER SELECTED DATES"] = means

    counts = wide.groupby(
        wide.index.to_period("M")
    ).count()

    counts.index = counts.index.astype(str)
    counts.index.name = "Month"
    counts.loc["TOTAL OBSERVATIONS"] = wide.count()

    tabs = st.tabs([
        "Daily prices",
        "Monthly averages",
        "Observation counts",
    ])

    for tab, frame, title in zip(
        tabs,
        [daily, monthly, counts],
        ["daily", "monthly", "counts"],
    ):
        with tab:
            if title == "counts":
                st.dataframe(
                    frame,
                    use_container_width=True,
                )
            else:
                st.dataframe(
                    frame.style.format(
                        "{:.4f}",
                        na_rep="",
                    ),
                    use_container_width=True,
                )

            if st.checkbox(
                "Prepare CSV download",
                key=key + "_" + title + "_prepare",
            ):
                st.download_button(
                    "Download CSV",
                    data=frame.to_csv().encode("utf-8-sig"),
                    file_name=f"{key}_{title}.csv",
                    mime="text/csv",
                    key=key + "_" + title + "_download",
                )

    st.caption(
        "The bottom average row uses daily numeric observations "
        "over the selected dates. Monthly averages use only "
        "dates inside that same range. Missing prices are "
        "excluded; zero and negative prices are included."
    )


st.title("Natural Gas Pricing Analysis")

section = st.sidebar.radio(
    "Dataset",
    options=[
        "ICE settlements",
        "NG Swing GDD master output",
    ],
)

if st.sidebar.button("Reload available data"):
    st.cache_data.clear()
    st.rerun()

st.sidebar.caption(
    "Reload clears cached data. Repository updates must "
    "first reach the deployed app."
)


# =================================================
# MASTER OUTPUT
# =================================================

if section == "NG Swing GDD master output":
    paths = sorted(
        (ROOT / "master").rglob("*.csv")
    )

    if not paths:
        st.error("No CSV files found in master.")
        st.stop()

    preferred = next(
        (
            path
            for path in paths
            if path.name.startswith("master_ng_swing_gdd")
        ),
        paths[0],
    )

    path = st.sidebar.selectbox(
        "Master file",
        paths,
        index=paths.index(preferred),
        format_func=lambda p: str(p.relative_to(ROOT)),
    )

    try:
        data = read_master(file_signature(path))
    except Exception as error:
        st.error(str(error))
        st.stop()

    st.header("NG Swing GDD master output")

    st.caption(
        f"Source: {path.relative_to(ROOT)}. "
        "The file does not distinguish actual from modeled "
        "values or identify historical run dates."
    )

    start, end = date_controls(
        data.index.min(),
        data.index.max(),
        "master_dates",
    )

    defaults = [
        hub
        for hub in ["HSC", "Henry", "Waha", "TETCO-M3"]
        if hub in data.columns
    ]

    hubs = st.sidebar.multiselect(
        "Master-output hubs",
        data.columns.tolist(),
        default=defaults or data.columns[:1].tolist(),
    )

    if not hubs:
        st.info("Select at least one hub.")
        st.stop()

    selected = data.loc[
        (data.index >= pd.Timestamp(start))
        & (data.index <= pd.Timestamp(end)),
        hubs,
    ]

    st.caption(
        "Each selected hub retains its own average."
    )

    show_report(
        selected,
        key="master",
        metric_label="Master output value",
    )


# =================================================
# ICE SETTLEMENTS
# =================================================

else:
    catalog_path = ARCHIVE / "catalog.parquet"
    manifest_path = ARCHIVE / "manifest.json"

    if (
        not catalog_path.exists()
        or not manifest_path.exists()
    ):
        st.info(
            "Run Prepare ICE Pricing Archive in GitHub first."
        )
        st.stop()

    try:
        catalog = read_catalog(
            file_signature(catalog_path)
        )
    except Exception as error:
        st.error(str(error))
        st.stop()

    if catalog.empty:
        st.warning("The instrument catalog is empty.")
        st.stop()

    st.header("ICE settlements")

    st.caption(
        f"{catalog['product'].nunique():,} products and "
        f"{catalog['contract'].nunique():,} ICE codes available."
    )

    st.sidebar.header("Products and hubs")

    st.sidebar.caption(
        "Empty filters include all available values. "
        "Select products and hubs before retrieving prices."
    )

    selected_catalog = multi_filter(
        catalog,
        "product",
        "Products",
        "ice_products",
    )

    selected_catalog = multi_filter(
        selected_catalog,
        "hub",
        "Hubs",
        "ice_hubs",
    )

    selected_catalog = multi_filter(
        selected_catalog,
        "contract",
        "ICE codes — optional",
        "ice_codes",
    )

    selected_catalog = multi_filter(
        selected_catalog,
        "contract_type",
        "Contract types — optional",
        "ice_types",
    )

    selected_catalog = multi_filter(
        selected_catalog,
        "strike",
        "Strikes — optional",
        "ice_strikes",
    )

    if selected_catalog.empty:
        st.info("No instruments match these filters.")
        st.stop()

    st.subheader("Trading-date range")

    # Bounds use the full archive, independently of
    # product, hub, and delivery-period selections.
    trade_start, trade_end = date_controls(
        catalog["first_trade_date"].min(),
        catalog["last_trade_date"].max(),
        "ice_trade_dates",
        "First trade date",
        "Last trade date",
    )

    st.subheader("Delivery-period range")

    delivery_start, delivery_end = date_controls(
        selected_catalog["strip"].min(),
        selected_catalog["strip"].max(),
        "ice_delivery_dates",
        "First delivery date",
        "Last delivery date",
    )

    selected_catalog = selected_catalog[
        selected_catalog["strip"].between(
            pd.Timestamp(delivery_start),
            pd.Timestamp(delivery_end),
        )
    ].copy()

    available_strips = sorted(
        selected_catalog["strip"]
        .dropna()
        .unique()
        .tolist()
    )

    specific_key = "ice_specific_strips"

    if specific_key in st.session_state:
        st.session_state[specific_key] = [
            value
            for value in st.session_state[specific_key]
            if value in available_strips
        ]

    specific_strips = st.multiselect(
        "Optional: select specific delivery dates within this range",
        options=available_strips,
        key=specific_key,
        format_func=lambda value: (
            pd.Timestamp(value).strftime("%b %d, %Y")
        ),
        help=(
            "Leave empty to include every available delivery "
            "date between the delivery start and end dates."
        ),
    )

    if specific_strips:
        selected_catalog = selected_catalog[
            selected_catalog["strip"].isin(specific_strips)
        ]

    if selected_catalog.empty:
        st.info(
            "No delivery strips match this range."
        )
        st.stop()

    metric_label = st.sidebar.selectbox(
        "Price field",
        ["Settlement price", "Net change"],
    )

    metric = (
        "settlement_price"
        if metric_label == "Settlement price"
        else "net_change"
    )

    st.caption(
        f"{len(selected_catalog):,} instrument identities selected | "
        f"{selected_catalog['product'].nunique():,} products | "
        f"{selected_catalog['hub'].nunique():,} hubs | "
        f"{selected_catalog['strip'].nunique():,} delivery dates."
    )

    # Display one contract family at a time to keep broad
    # multi-product selections responsive.
    families = (
        selected_catalog[FAMILY]
        .drop_duplicates()
        .reset_index(drop=True)
    )

    family_labels = [
        " | ".join(
            f"{column}: {value}"
            for column, value in zip(FAMILY, row)
            if pd.notna(value) and str(value) != ""
        )
        for row in families.itertuples(
            index=False,
            name=None,
        )
    ]

    # An explicit button prevents querying the archive
    # after every intermediate filter edit.
    selection_signature = (
        file_signature(manifest_path),
        tuple(
            instrument_tuple(row)
            for row in selected_catalog[
                IDENTITY
            ].itertuples(index=False, name=None)
        ),
        str(trade_start),
        str(trade_end),
    )

    if st.button(
        "Apply filters and load prices",
        type="primary",
    ):
        st.session_state["active_ice_selection"] = (
            selection_signature
        )

    if (
        st.session_state.get("active_ice_selection")
        != selection_signature
    ):
        st.info(
            "Set your filters, then click "
            "Apply filters and load prices."
        )
        st.stop()

    family_position = st.selectbox(
        "Product / hub / contract group to display",
        options=list(range(len(families))),
        format_func=lambda position: family_labels[position],
        help=(
            "All selected groups are available here. "
            "Each group has its own prices, strip average, "
            "and average rows."
        ),
    )

    family_row = families.iloc[family_position]
    family_catalog = selected_catalog.copy()

    for column in FAMILY:
        value = family_row[column]

        if pd.isna(value):
            family_catalog = family_catalog[
                family_catalog[column].isna()
            ]
        else:
            family_catalog = family_catalog[
                family_catalog[column].eq(value)
            ]

    instruments = tuple(
        instrument_tuple(row)
        for row in family_catalog[
            IDENTITY
        ].drop_duplicates().itertuples(
            index=False,
            name=None,
        )
    )

    with st.spinner("Retrieving selected pricing..."):
        try:
            data = query_prices(
                file_signature(manifest_path),
                instruments,
                trade_start,
                trade_end,
            )
        except Exception as error:
            st.error(str(error))
            st.stop()

    if data.empty:
        st.info(
            "This group has no records within the "
            "selected trading-date range."
        )
        st.stop()

    records = data.drop_duplicates(
        ["trade_date"] + IDENTITY
    ).copy()

    st.subheader(
        family_labels[family_position]
    )

    # Use the catalog, not only returned prices, to retain
    # delivery columns with no observations in this date range.
    expiry_counts = (
        family_catalog.groupby("strip")[
            "expiration_date"
        ]
        .nunique(dropna=False)
    )

    multiple_expirations = expiry_counts.gt(1).any()

    def delivery_labels(frame):
        labels = frame["strip"].dt.strftime("%Y-%m-%d")

        if multiple_expirations:
            labels = (
                labels
                + " | expires "
                + frame["expiration_date"]
                .dt.strftime("%Y-%m-%d")
                .fillna("unspecified")
            )

        return labels

    records["delivery_column"] = delivery_labels(records)

    expected_columns = sorted(
        delivery_labels(family_catalog).unique().tolist()
    )

    if records.duplicated(
        ["trade_date", "delivery_column"]
    ).any():
        st.error(
            "Multiple instruments map to one price-table cell."
        )
        st.stop()

    wide = records.pivot(
        index="trade_date",
        columns="delivery_column",
        values=metric,
    )

    wide.columns.name = None

    st.caption(
        f"{len(records):,} unique pricing records. "
        "Strip averages equally weight the displayed delivery "
        "columns. Products, hubs, contract types, and strikes "
        "are kept separate."
    )

    show_report(
        wide,
        key="ice_report_" + str(family_position),
        metric_label=metric_label,
        strip_average=True,
        expected_columns=expected_columns,
    )

    with st.expander(
        "Underlying records and source files"
    ):
        st.dataframe(
            data,
            use_container_width=True,
            hide_index=True,
        )

        if st.checkbox(
            "Prepare underlying-records CSV",
            key="ice_raw_prepare",
        ):
            st.download_button(
                "Download underlying records",
                data=data.to_csv(
                    index=False
                ).encode("utf-8-sig"),
                file_name="ice_underlying_records.csv",
                mime="text/csv",
            )
