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
    df.columns = [
        str(column).strip()
        for column in df.columns
    ]

    date_column = next(
        (
            column
            for column in df.columns
            if column.lower() == "date"
        ),
        None,
    )

    if date_column is None:
        raise ValueError(
            "The master CSV has no Date column."
        )

    df = df.rename(
        columns={date_column: "Date"}
    )

    df["Date"] = pd.to_datetime(
        df["Date"],
        errors="raise",
    )

    if df["Date"].isna().any():
        raise ValueError(
            "The master CSV contains missing dates."
        )

    for column in df.columns:
        if column != "Date":
            df[column] = pd.to_numeric(
                df[column],
                errors="raise",
            )

    df = df.drop_duplicates()

    if df["Date"].duplicated().any():
        raise ValueError(
            "The master CSV has conflicting "
            "records for the same date."
        )

    return (
        df.sort_values("Date")
        .set_index("Date")
    )


@st.cache_data(max_entries=4, show_spinner=False)
def query_prices(
    archive_signature,
    contract,
    product,
    hub,
    start,
    end,
):
    connection = duckdb.connect()

    try:
        connection.execute(
            "SET memory_limit='512MB'"
        )

        connection.execute(
            "SET threads=2"
        )

        connection.read_parquet(
            str(
                ARCHIVE
                / "parts"
                / "*.parquet"
            ),
            union_by_name=True,
        ).create_view("prices")

        # Retrieve only the selected product, hub,
        # code, and date range.
        result = connection.execute(
            """
            SELECT DISTINCT *
            FROM prices
            WHERE contract = ?
              AND product = ?
              AND hub = ?
              AND trade_date >= CAST(? AS DATE)
              AND trade_date <= CAST(? AS DATE)
            """,
            [
                contract,
                product,
                hub,
                str(start),
                str(end),
            ],
        ).df()

        return result

    finally:
        connection.close()


def date_controls(first, last, key):
    first = pd.Timestamp(first).date()
    last = pd.Timestamp(last).date()

    left, right = st.columns(2)

    start = left.date_input(
        "Start date",
        value=first,
        min_value=first,
        max_value=last,
        key=key + "_start",
    )

    end = right.date_input(
        "End date",
        value=last,
        min_value=first,
        max_value=last,
        key=key + "_end",
    )

    if start > end:
        st.error(
            "Start date must be on or before end date."
        )
        st.stop()

    return start, end


def filter_catalog(
    frame,
    column,
    label,
    key,
):
    options = sorted(
        frame[column]
        .dropna()
        .unique()
        .tolist()
    )

    if not options:
        return frame

    selected = st.sidebar.multiselect(
        label,
        options,
        key=key,
        help=(
            "Leave empty to include every "
            "available value."
        ),
    )

    if selected:
        return frame[
            frame[column].isin(selected)
        ]

    return frame


def report(
    wide,
    key,
    strip_average=False,
    metric_label="Price",
):
    wide = wide.sort_index().copy()

    if wide.empty:
        st.info(
            "No values match these dates."
        )
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
            help=(
                "When checked, the strip average "
                "is blank on a date if any selected "
                "delivery column has no price."
            ),
        )

        if include_average:
            average = wide.mean(axis=1)

            if require_all:
                average = average.where(
                    wide.notna().all(axis=1)
                )

            wide[
                "Selected-strip average"
            ] = average

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

    if frequency == "Daily":
        chart = wide.copy()
    else:
        chart = monthly.copy()
        chart.index = (
            chart.index.to_timestamp()
        )

    if chart.shape[1] <= 30:
        chart.index.name = "Date"

        long = (
            chart.reset_index()
            .melt(
                id_vars="Date",
                var_name="Series",
                value_name="Value",
            )
        )

        figure = px.line(
            long,
            x="Date",
            y="Value",
            color="Series",
            markers=True,
            labels={
                "Value": metric_label,
            },
        )

        figure.update_traces(
            connectgaps=False
        )

        for trace in figure.data:
            if (
                trace.name
                == "Selected-strip average"
            ):
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
            "Select 30 or fewer series to chart "
            "them. All selected series remain "
            "available in the tables."
        )

    daily = wide.copy()
    daily.index = (
        daily.index.strftime("%Y-%m-%d")
    )
    daily.index.name = "Date"

    daily.loc[
        "AVERAGE OVER SELECTED DATES"
    ] = means

    monthly.index = (
        monthly.index.astype(str)
    )
    monthly.index.name = "Month"

    monthly.loc[
        "AVERAGE OVER SELECTED DATES"
    ] = means

    counts = wide.groupby(
        wide.index.to_period("M")
    ).count()

    counts.index = (
        counts.index.astype(str)
    )
    counts.index.name = "Month"

    counts.loc[
        "TOTAL OBSERVATIONS"
    ] = wide.count()

    tabs = st.tabs(
        [
            "Daily prices",
            "Monthly averages",
            "Observation counts",
        ]
    )

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

            prepare_download = st.checkbox(
                "Prepare CSV download",
                key=(
                    key
                    + "_"
                    + title
                    + "_prepare"
                ),
            )

            if prepare_download:
                st.download_button(
                    "Download CSV",
                    data=frame.to_csv().encode(
                        "utf-8-sig"
                    ),
                    file_name=(
                        f"{key}_{title}.csv"
                    ),
                    mime="text/csv",
                    key=(
                        key
                        + "_"
                        + title
                        + "_download"
                    ),
                )

    st.caption(
        "The bottom average row uses daily numeric "
        "observations within the selected dates, "
        "not an equal-weight mean of monthly "
        "averages. Missing values are excluded; "
        "zero and negative prices are included."
    )


st.title(
    "Natural Gas Pricing Analysis"
)

section = st.sidebar.radio(
    "Dataset",
    [
        "ICE settlements",
        "NG Swing GDD master output",
    ],
)

if st.sidebar.button(
    "Reload available data"
):
    st.cache_data.clear()
    st.rerun()

st.sidebar.caption(
    "Reload clears this app's cache. "
    "GitHub updates must first reach "
    "the deployed app."
)


# -------------------------------------------------
# NG SWING GDD MASTER OUTPUT
# -------------------------------------------------

if section == "NG Swing GDD master output":
    paths = sorted(
        (ROOT / "master").rglob("*.csv")
    )

    if not paths:
        st.error(
            "No CSV files found in master."
        )
        st.stop()

    preferred = next(
        (
            path
            for path in paths
            if path.name.startswith(
                "master_ng_swing_gdd"
            )
        ),
        paths[0],
    )

    path = st.sidebar.selectbox(
        "Master file",
        options=paths,
        index=paths.index(preferred),
        format_func=lambda selected_path: str(
            selected_path.relative_to(ROOT)
        ),
    )

    try:
        dataframe = read_master(
            file_signature(path)
        )

    except Exception as error:
        st.error(str(error))
        st.stop()

    st.header(
        "NG Swing GDD master output"
    )

    st.caption(
        "Source: "
        + str(path.relative_to(ROOT))
        + ". Values may include modeled dates; "
        "the file does not identify actual "
        "versus modeled records."
    )

    start, end = date_controls(
        dataframe.index.min(),
        dataframe.index.max(),
        "master",
    )

    defaults = [
        column
        for column in [
            "HSC",
            "Henry",
            "Waha",
            "TETCO-M3",
        ]
        if column in dataframe.columns
    ]

    hubs = st.sidebar.multiselect(
        "Hubs",
        options=dataframe.columns.tolist(),
        default=(
            defaults
            or dataframe.columns[:1].tolist()
        ),
    )

    if not hubs:
        st.info(
            "Select at least one hub."
        )
        st.stop()

    selected = dataframe.loc[
        (
            dataframe.index
            >= pd.Timestamp(start)
        )
        & (
            dataframe.index
            <= pd.Timestamp(end)
        ),
        hubs,
    ]

    st.caption(
        "Each hub has its own selected-period "
        "average. Different hubs are not "
        "blended into one price."
    )

    report(
        selected,
        "master",
        metric_label="Master output value",
    )


# -------------------------------------------------
# ICE SETTLEMENTS
# -------------------------------------------------

else:
    catalog_path = (
        ARCHIVE / "catalog.parquet"
    )

    manifest_path = (
        ARCHIVE / "manifest.json"
    )

    if (
        not catalog_path.exists()
        or not manifest_path.exists()
    ):
        st.info(
            "Run the Prepare ICE Pricing Archive "
            "workflow in GitHub first."
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
        st.warning(
            "The ICE instrument catalog is empty."
        )
        st.stop()

    st.header(
        "ICE settlements"
    )

    st.caption(
        f"{catalog['product'].nunique():,} products "
        f"and {catalog['contract'].nunique():,} "
        "ICE codes available. Choose a product "
        "first, then its hub and matching code."
    )

    st.sidebar.header(
        "Select product and hub"
    )

    # Product is the first selection.
    products = sorted(
        catalog["product"]
        .dropna()
        .unique()
        .tolist()
    )

    product = st.sidebar.selectbox(
        "Product",
        options=products,
        key="product_first",
        help=(
            "Type part of the product description "
            "to search the dropdown."
        ),
    )

    product_catalog = catalog[
        catalog["product"].eq(product)
    ].copy()

    # Hub is filtered by the chosen product.
    hubs = sorted(
        product_catalog["hub"]
        .dropna()
        .unique()
        .tolist()
    )

    hub = st.sidebar.selectbox(
        "Hub",
        options=hubs,
        key="product_first_hub",
    )

    hub_catalog = product_catalog[
        product_catalog["hub"].eq(hub)
    ].copy()

    # Only matching codes are offered.
    codes = sorted(
        hub_catalog["contract"]
        .dropna()
        .unique()
        .tolist()
    )

    code = st.sidebar.selectbox(
        "Matching ICE code",
        options=codes,
        key="product_first_code",
        help=(
            "Only codes associated with the "
            "selected product and hub appear."
        ),
    )

    code_catalog = hub_catalog[
        hub_catalog["contract"].eq(code)
    ].copy()

    selected_catalog = code_catalog.copy()

    selected_catalog = filter_catalog(
        selected_catalog,
        "contract_type",
        "Contract type",
        key="ice_type",
    )

    selected_catalog = filter_catalog(
        selected_catalog,
        "strike",
        "Strike, if applicable",
        key="ice_strike",
    )

    strips = sorted(
        selected_catalog["strip"]
        .dropna()
        .unique()
        .tolist()
    )

    chosen_strips = st.sidebar.multiselect(
        "Delivery strips",
        options=strips,
        default=strips[:1],
        format_func=lambda date: (
            pd.Timestamp(date).strftime(
                "%b %d, %Y"
            )
        ),
        help=(
            "Select one or more delivery periods. "
            "Clear the selection to include "
            "all available strips."
        ),
    )

    if chosen_strips:
        selected_catalog = selected_catalog[
            selected_catalog["strip"].isin(
                chosen_strips
            )
        ]

    if selected_catalog.empty:
        st.info(
            "No instruments match these filters."
        )
        st.stop()

    # Stable date bounds for the selected
    # product, hub, and code.
    date_key = (
        "ice_dates_"
        + str(
            products.index(product)
        )
        + "_"
        + str(hubs.index(hub))
        + "_"
        + code
    )

    start, end = date_controls(
        code_catalog[
            "first_trade_date"
        ].min(),
        code_catalog[
            "last_trade_date"
        ].max(),
        date_key,
    )

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

    st.caption(
        f"Product: {product} | "
        f"Hub: {hub} | "
        f"ICE code: {code}"
    )

    with st.spinner(
        "Retrieving selected pricing..."
    ):
        try:
            data = query_prices(
                file_signature(manifest_path),
                code,
                product,
                hub,
                start,
                end,
            )

        except Exception as error:
            st.error(str(error))
            st.stop()

    data = data.merge(
        selected_catalog[
            IDENTITY
        ].drop_duplicates(),
        on=IDENTITY,
        how="inner",
        validate="many_to_one",
    )

    if data.empty:
        st.info(
            "No prices match these dates "
            "and instruments."
        )
        st.stop()

    # Identical records in multiple source files
    # count once in the pricing calculations.
    # The archive builder rejects conflicts.
    records = data.drop_duplicates(
        ["trade_date"] + IDENTITY
    )

    st.caption(
        f"{len(records):,} unique matching "
        "pricing records."
    )

    groups = records.groupby(
        FAMILY,
        dropna=False,
        observed=True,
        sort=True,
    )

    if groups.ngroups > 20:
        st.info(
            "More than 20 contract groups "
            "selected. Narrow contract type "
            "or strike filters."
        )

    else:
        for position, (
            labels,
            group,
        ) in enumerate(groups):
            if not isinstance(
                labels,
                tuple,
            ):
                labels = (labels,)

            title = " | ".join(
                f"{name}: {value}"
                for name, value in zip(
                    FAMILY,
                    labels,
                )
                if (
                    pd.notna(value)
                    and str(value) != ""
                )
            )

            st.subheader(title)

            group = group.copy()

            multiple_expirations = (
                group.groupby("strip")[
                    "expiration_date"
                ]
                .nunique(dropna=False)
                .gt(1)
                .any()
            )

            group["column"] = (
                group["strip"]
                .dt.strftime("%Y-%m-%d")
            )

            if multiple_expirations:
                group["column"] += (
                    " | expires "
                    + group[
                        "expiration_date"
                    ]
                    .dt.strftime("%Y-%m-%d")
                    .fillna("unspecified")
                )

            if group.duplicated(
                [
                    "trade_date",
                    "column",
                ]
            ).any():
                st.error(
                    "Multiple instrument identities "
                    "map to one table cell. "
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
                + ". Selected-strip averages "
                "equally weight the displayed "
                "delivery columns. Different "
                "products, hubs, contract types, "
                "and strikes remain separate."
            )

            report(
                wide,
                key=f"ice_{position}",
                strip_average=True,
                metric_label=metric_label,
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
            key="underlying_prepare",
        ):
            st.download_button(
                "Download underlying records",
                data=data.to_csv(
                    index=False
                ).encode("utf-8-sig"),
                file_name=(
                    "ice_underlying_records.csv"
                ),
                mime="text/csv",
            )
