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
    return (str(path), stat.st_mtime_ns, stat.st_size)


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
        raise ValueError("Master CSV has no Date column.")

    df = df.rename(columns={date_column: "Date"})
    df["Date"] = pd.to_datetime(df["Date"], errors="raise")

    for column in df.columns:
        if column != "Date":
            df[column] = pd.to_numeric(
                df[column],
                errors="raise",
            )

    df = df.drop_duplicates()

    if df["Date"].duplicated().any():
        raise ValueError(
            "Master CSV has conflicting records for one date."
        )

    return df.sort_values("Date").set_index("Date")


@st.cache_data(max_entries=4, show_spinner=False)
def query_prices(archive_signature, contract, start, end):
    con = duckdb.connect()

    try:
        con.execute("SET memory_limit='512MB'")
        con.execute("SET threads=2")

        con.read_parquet(
            str(ARCHIVE / "parts" / "*.parquet"),
            union_by_name=True,
        ).create_view("prices")

        # Only the requested code and date range enter Pandas.
        result = con.execute(
            """
            SELECT DISTINCT *
            FROM prices
            WHERE contract = ?
              AND trade_date >= CAST(? AS DATE)
              AND trade_date <= CAST(? AS DATE)
            """,
            [contract, str(start), str(end)],
        ).df()

        return result

    finally:
        con.close()


def date_controls(first, last, key):
    first = pd.Timestamp(first).date()
    last = pd.Timestamp(last).date()

    left, right = st.columns(2)

    start = left.date_input(
        "Start date",
        first,
        min_value=first,
        max_value=last,
        key=key + "_start",
    )

    end = right.date_input(
        "End date",
        last,
        min_value=first,
        max_value=last,
        key=key + "_end",
    )

    if start > end:
        st.error("Start date must be on or before end date.")
        st.stop()

    return start, end


def filter_catalog(frame, column, label):
    options = sorted(
        frame[column].dropna().unique().tolist()
    )

    selected = st.sidebar.multiselect(
        label,
        options,
        key="filter_" + column,
        help="Empty means all available values.",
    )

    if selected:
        return frame[frame[column].isin(selected)]

    return frame


def report(wide, key, strip_average=False):
    wide = wide.sort_index().copy()

    if wide.empty:
        st.info("No values match these dates.")
        return

    if strip_average:
        include_average = st.checkbox(
            "Show selected-strip average",
            value=True,
            key=key + "_average",
        )

        require_all = st.checkbox(
            "Require all selected strips for the average",
            value=True,
            key=key + "_complete",
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
        ["Daily", "Monthly"],
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
            value_name="Price",
        )

        fig = px.line(
            long,
            x="Date",
            y="Price",
            color="Series",
            markers=True,
        )

        fig.update_traces(connectgaps=False)

        for trace in fig.data:
            if trace.name == "Selected-strip average":
                trace.line.update(
                    color="black",
                    width=4,
                    dash="dash",
                )

        fig.update_layout(
            height=520,
            legend={
                "orientation": "h",
                "y": -0.25,
            },
        )

        st.plotly_chart(fig, use_container_width=True)

    else:
        st.info(
            "Select 30 or fewer series to chart them. "
            "The tables include all selected series."
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
            st.dataframe(
                frame,
                use_container_width=True,
            )

            # Prepare downloads only when requested.
            if st.checkbox(
                "Prepare CSV download",
                key=key + "_" + title + "_prepare",
            ):
                st.download_button(
                    "Download CSV",
                    frame.to_csv().encode("utf-8-sig"),
                    f"{key}_{title}.csv",
                    "text/csv",
                    key=key + "_" + title + "_download",
                )

    st.caption(
        "The bottom average row uses daily numeric observations "
        "within the selected dates, not an equal-weight mean "
        "of monthly averages. Missing values are excluded; "
        "zero and negative prices are included."
    )


st.title("Natural Gas Pricing Analysis")

section = st.sidebar.radio(
    "Dataset",
    [
        "ICE settlements",
        "NG Swing GDD master output",
    ],
)

if st.sidebar.button("Reload available data"):
    st.cache_data.clear()
    st.rerun()

st.sidebar.caption(
    "Reload clears this app's cache. GitHub updates must "
    "first reach the deployed app."
)


if section == "NG Swing GDD master output":
    paths = sorted((ROOT / "master").rglob("*.csv"))

    if not paths:
        st.error("No CSV files found in master.")
        st.stop()

    preferred = next(
        (
            p for p in paths
            if p.name.startswith("master_ng_swing_gdd")
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
        df = read_master(file_signature(path))
    except Exception as error:
        st.error(str(error))
        st.stop()

    st.header("NG Swing GDD master output")

    st.caption(
        "Source: " + str(path.relative_to(ROOT))
        + ". Values may include modeled dates; the file "
        "does not identify actual versus modeled records."
    )

    start, end = date_controls(
        df.index.min(),
        df.index.max(),
        "master",
    )

    defaults = [
        c for c in ["HSC", "Henry", "Waha", "TETCO-M3"]
        if c in df.columns
    ]

    hubs = st.sidebar.multiselect(
        "Hubs",
        df.columns.tolist(),
        default=defaults or df.columns[:1].tolist(),
    )

    if not hubs:
        st.info("Select at least one hub.")
        st.stop()

    selected = df.loc[
        (df.index >= pd.Timestamp(start))
        & (df.index <= pd.Timestamp(end)),
        hubs,
    ]

    report(selected, "master")


else:
    catalog_path = ARCHIVE / "catalog.parquet"
    manifest_path = ARCHIVE / "manifest.json"

    if not catalog_path.exists() or not manifest_path.exists():
        st.info(
            "Run the Prepare ICE Pricing Archive workflow "
            "in GitHub first."
        )
        st.stop()

    catalog = read_catalog(file_signature(catalog_path))

    st.header("ICE settlements")

    st.caption(
        f"{catalog['contract'].nunique():,} ICE codes available. "
        "Dropdowns use the instrument catalog; pricing is "
        "queried only for your selected code and dates."
    )

    codes = sorted(catalog["contract"].unique().tolist())

    preferred = next(
        (c for c in ["HXS", "FQT", "TPH"] if c in codes),
        codes[0],
    )

    code = st.sidebar.selectbox(
        "ICE code",
        codes,
        index=codes.index(preferred),
    )

    selected_catalog = catalog[
        catalog["contract"].eq(code)
    ].copy()

    for column, label in [
        ("product", "Product"),
        ("hub", "Hub"),
        ("contract_type", "Contract type"),
        ("strike", "Strike, if applicable"),
    ]:
        selected_catalog = filter_catalog(
            selected_catalog,
            column,
            label,
        )

    strips = sorted(
        selected_catalog["strip"].dropna().unique().tolist()
    )

    chosen_strips = st.sidebar.multiselect(
        "Delivery strips",
        strips,
        default=strips[:1],
        format_func=lambda d: (
            pd.Timestamp(d).strftime("%b %d, %Y")
        ),
        help="Empty includes all available strips.",
    )

    if chosen_strips:
        selected_catalog = selected_catalog[
            selected_catalog["strip"].isin(chosen_strips)
        ]

    if selected_catalog.empty:
        st.info("No instruments match these filters.")
        st.stop()

    # Use the full code's coverage to keep date controls stable
    # when delivery-strip filters change.
    coverage = catalog[catalog["contract"].eq(code)]

    start, end = date_controls(
        coverage["first_trade_date"].min(),
        coverage["last_trade_date"].max(),
        "ice_" + code,
    )

    metric_label = st.sidebar.selectbox(
        "Price field",
        ["Settlement price", "Net change"],
    )

    metric = (
        "settlement_price"
        if metric_label == "Settlement price"
        else "net_change"
    )

    with st.spinner("Retrieving selected pricing..."):
        try:
            data = query_prices(
                file_signature(manifest_path),
                code,
                start,
                end,
            )
        except Exception as error:
            st.error(str(error))
            st.stop()

    # Catalog identities are unique. Null keys match in Pandas.
    data = data.merge(
        selected_catalog[IDENTITY].drop_duplicates(),
        on=IDENTITY,
        how="inner",
        validate="many_to_one",
    )

    if data.empty:
        st.info("No prices match these dates and instruments.")
        st.stop()

    # Remove identical prices repeated in different source files.
    # Conflicting prices were rejected by the archive builder.
    records = data.drop_duplicates(
        ["trade_date"] + IDENTITY
    )

    st.caption(
        f"{len(records):,} unique matching pricing records."
    )

    groups = records.groupby(
        FAMILY,
        dropna=False,
        observed=True,
        sort=True,
    )

    if groups.ngroups > 20:
        st.info(
            "More than 20 contract groups selected. "
            "Narrow product, hub, type, or strike filters."
        )

    else:
        for position, (labels, group) in enumerate(groups):
            if not isinstance(labels, tuple):
                labels = (labels,)

            title = " | ".join(
                f"{name}: {value}"
                for name, value in zip(FAMILY, labels)
                if pd.notna(value) and str(value) != ""
            )

            st.subheader(title)
            group = group.copy()

            # Expiration differentiates instruments only when
            # a delivery strip has more than one expiration.
            multiple_expirations = (
                group.groupby("strip")["expiration_date"]
                .nunique(dropna=False)
                .gt(1)
                .any()
            )

            group["column"] = group["strip"].dt.strftime(
                "%Y-%m-%d"
            )

            if multiple_expirations:
                group["column"] += (
                    " | expires "
                    + group["expiration_date"]
                    .dt.strftime("%Y-%m-%d")
                    .fillna("unspecified")
                )

            if group.duplicated(
                ["trade_date", "column"]
            ).any():
                st.error(
                    "Multiple identities map to one table cell. "
                    "Narrow the filters."
                )
                continue

            wide = group.pivot(
                index="trade_date",
                columns="column",
                values=metric,
            )

            wide.columns.name = None

            st.caption(
                metric_label
                + ". Selected-strip averages equally weight "
                "the displayed delivery columns; different "
                "hubs and products remain separate."
            )

            report(
                wide,
                f"ice_{position}",
                strip_average=True,
            )

    with st.expander("Underlying records and source files"):
        st.dataframe(
            data,
            use_container_width=True,
            hide_index=True,
        )
