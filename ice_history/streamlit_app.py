from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(
    page_title="Natural Gas Pricing History",
    layout="wide",
)

ROOT = Path(__file__).resolve().parents[1]

ICE_ID = [
    "contract",
    "product",
    "hub",
    "strip",
    "contract_type",
    "strike",
    "expiration_date",
    "product_id",
]

ICE_FAMILY = [
    "contract",
    "product",
    "hub",
    "contract_type",
    "strike",
    "expiration_date",
    "product_id",
]

ICE_COLUMNS = [
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


def manifest(paths):
    return tuple(
        (
            str(path),
            path.stat().st_mtime_ns,
            path.stat().st_size,
        )
        for path in sorted(paths)
    )


@st.cache_data(show_spinner=False, max_entries=1)
def load_ice(file_manifest, codes):
    frames = []

    for filename, modified, size in file_manifest:
        df = pd.read_excel(
            filename,
            engine="openpyxl",
            dtype=object,
        )

        df.columns = [
            str(c).strip().lower().replace(" ", "_")
            for c in df.columns
        ]

        missing = set(ICE_COLUMNS) - set(df.columns)

        if missing:
            raise ValueError(
                f"{Path(filename).name}: missing columns "
                f"{sorted(missing)}"
            )

        df = df[ICE_COLUMNS].dropna(how="all").copy()

        for column in [
            "contract",
            "product",
            "hub",
            "contract_type",
            "product_id",
        ]:
            df[column] = (
                df[column].fillna("").astype(str).str.strip()
            )

        df["contract"] = df["contract"].str.upper()

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

            if (nonblank & parsed.isna()).any():
                raise ValueError(
                    f"{Path(filename).name}: invalid {column}"
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

            if (nonblank & parsed.isna()).any():
                raise ValueError(
                    f"{Path(filename).name}: invalid {column}"
                )

            df[column] = parsed

        if (
            df["trade_date"].isna()
            | df["strip"].isna()
            | df["contract"].eq("")
        ).any():
            raise ValueError(
                f"{Path(filename).name}: missing trade date, "
                "strip, or ICE code"
            )

        df["source_file"] = str(
            Path(filename).relative_to(ROOT)
        )

        frames.append(df)

    if not frames:
        raise ValueError("No matching ICE records were found.")

    data = pd.concat(frames, ignore_index=True)
    del frames

    key = ["trade_date"] + ICE_ID
    duplicated = data.duplicated(key, keep=False)

    if duplicated.any():
        duplicate_rows = data.loc[duplicated]

        counts = (
            duplicate_rows.groupby(
                key,
                dropna=False,
                observed=True,
            )[["settlement_price", "net_change"]]
            .nunique(dropna=False)
        )

        conflicts = counts[(counts > 1).any(axis=1)]

        if not conflicts.empty:
            examples = (
                conflicts.reset_index()[
                    ["trade_date", "contract", "hub", "strip"]
                ]
                .head(5)
                .to_string(index=False)
            )

            raise ValueError(
                f"{len(conflicts)} conflicting duplicate keys. "
                "Resolve the source records first.\n\n"
                + examples
            )

        sources = (
            duplicate_rows.groupby(
                key,
                dropna=False,
                observed=True,
            )["source_file"]
            .agg(lambda s: "; ".join(sorted(set(s))))
            .reset_index()
        )

        clean_duplicates = (
            duplicate_rows.drop(columns="source_file")
            .drop_duplicates(key)
            .merge(
                sources,
                on=key,
                how="left",
                validate="one_to_one",
            )
        )

        data = pd.concat(
            [
                data.loc[~duplicated],
                clean_duplicates,
            ],
            ignore_index=True,
        )

    data = data.sort_values(
        ["contract", "hub", "strip", "trade_date"]
    ).reset_index(drop=True)

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


@st.cache_data(show_spinner=False, max_entries=2)
def load_master(filename, modified, size):
    df = pd.read_csv(filename)
    df.columns = [str(c).strip() for c in df.columns]

    date_column = next(
        (c for c in df.columns if c.lower() == "date"),
        None,
    )

    if date_column is None:
        raise ValueError("The master file needs a Date column.")

    df = df.rename(columns={date_column: "Date"})
    original_dates = df["Date"]
    df["Date"] = pd.to_datetime(
        original_dates,
        errors="coerce",
    )

    if df["Date"].isna().any():
        raise ValueError("The master file contains invalid dates.")

    for column in df.columns:
        if column == "Date":
            continue

        original = df[column]
        parsed = pd.to_numeric(original, errors="coerce")

        nonblank = (
            original.notna()
            & original.astype(str).str.strip().ne("")
        )

        if (nonblank & parsed.isna()).any():
            raise ValueError(
                f"Master column {column} contains nonnumeric values."
            )

        df[column] = parsed

    # Exact duplicate rows are harmless; conflicting dates are not.
    df = df.drop_duplicates()

    if df["Date"].duplicated().any():
        raise ValueError(
            "The master file contains different records "
            "for the same date."
        )

    return df.sort_values("Date").set_index("Date")


def multiselect_filter(frame, column, label, key, default=None):
    options = sorted(
        frame[column].dropna().unique().tolist()
    )

    selected = st.sidebar.multiselect(
        label,
        options,
        default=default or [],
        key=key,
        help="An empty selection includes every available value.",
    )

    if selected:
        return frame[frame[column].isin(selected)]

    return frame


def select_dates(index, key):
    first = pd.Timestamp(index.min()).date()
    last = pd.Timestamp(index.max()).date()

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
        st.error("Start date must be on or before end date.")
        st.stop()

    return pd.Timestamp(start), pd.Timestamp(end)


def download(frame, label, filename, key):
    st.download_button(
        label,
        frame.to_csv(index=True).encode("utf-8-sig"),
        filename,
        "text/csv",
        key=key,
    )


def price_report(wide, key, selected_strip_average=False):
    """
    Each column is a distinct series.

    Daily table: bottom row is the selected-date mean.
    Monthly table: bottom row is the same daily-observation mean,
    not an unweighted mean of unequal-length months.
    """
    wide = wide.sort_index().copy()
    wide = wide.apply(pd.to_numeric, errors="coerce")

    if wide.empty or wide.shape[1] == 0:
        st.info("No prices match this selection.")
        return

    original_columns = wide.columns.tolist()

    include_average = False

    if selected_strip_average:
        include_average = st.checkbox(
            "Add selected-strip average",
            value=True,
            key=key + "_include_average",
        )

        require_complete = st.checkbox(
            "Require a price for every selected strip",
            value=True,
            key=key + "_complete",
            help=(
                "When checked, a daily strip average is blank "
                "if any selected strip is missing a price."
            ),
        )

        if include_average:
            average = wide[original_columns].mean(axis=1)

            if require_complete:
                complete = wide[original_columns].notna().all(axis=1)
                average = average.where(complete)

            wide["Selected-strip average"] = average

    selected_means = wide.mean(axis=0)

    daily = wide.copy()
    daily.index = daily.index.strftime("%Y-%m-%d")
    daily.index.name = "Date"

    daily.loc[
        "AVERAGE OVER SELECTED DATES"
    ] = selected_means

    monthly = wide.groupby(
        wide.index.to_period("M")
    ).mean()

    monthly.index = monthly.index.astype(str)
    monthly.index.name = "Month"

    monthly.loc[
        "AVERAGE OVER SELECTED DATES"
    ] = selected_means

    counts = wide.groupby(
        wide.index.to_period("M")
    ).count()

    counts.index = counts.index.astype(str)
    counts.index.name = "Month"

    counts.loc[
        "TOTAL OBSERVATIONS"
    ] = wide.count()

    chart_type = st.radio(
        "Chart frequency",
        ["Daily", "Monthly average"],
        horizontal=True,
        key=key + "_frequency",
    )

    chart = (
        wide.copy()
        if chart_type == "Daily"
        else wide.groupby(
            wide.index.to_period("M")
        ).mean()
    )

    if chart_type != "Daily":
        chart.index = chart.index.to_timestamp()

    if chart.shape[1] > 30:
        st.info(
            "Select 30 or fewer series to show the chart. "
            "All selected series remain in the tables."
        )

    else:
        chart.index.name = "Date"

        long = (
            chart.reset_index()
            .melt(
                id_vars="Date",
                var_name="Series",
                value_name="Value",
            )
        )

        fig = px.line(
            long,
            x="Date",
            y="Value",
            color="Series",
            markers=True,
        )

        fig.update_traces(connectgaps=False)

        if include_average:
            for trace in fig.data:
                if trace.name == "Selected-strip average":
                    trace.line.update(
                        color="black",
                        width=4,
                        dash="dash",
                    )

        fig.update_layout(
            height=530,
            legend={
                "orientation": "h",
                "y": -0.25,
                "yanchor": "top",
            },
        )

        st.plotly_chart(
            fig,
            use_container_width=True,
        )

    daily_tab, monthly_tab, counts_tab = st.tabs(
        [
            "Daily table",
            "Monthly averages",
            "Observation counts",
        ]
    )

    with daily_tab:
        st.dataframe(
            daily.style.format("{:.4f}", na_rep=""),
            use_container_width=True,
        )

        download(
            daily,
            "Download daily table",
            key + "_daily.csv",
            key + "_daily_download",
        )

    with monthly_tab:
        st.dataframe(
            monthly.style.format("{:.4f}", na_rep=""),
            use_container_width=True,
        )

        st.caption(
            "The bottom row averages the underlying daily "
            "observations across the selected dates. "
            "It is not an equal-weight average of monthly averages."
        )

        download(
            monthly,
            "Download monthly table",
            key + "_monthly.csv",
            key + "_monthly_download",
        )

    with counts_tab:
        st.dataframe(
            counts,
            use_container_width=True,
        )


st.title("Natural Gas Pricing Analysis")

section = st.sidebar.radio(
    "Dataset",
    [
        "ICE settlements",
        "NG Swing GDD master output",
    ],
)

if st.sidebar.button("Reload available files"):
    st.cache_data.clear()
    st.rerun()

st.sidebar.caption(
    "Reload clears cached data on this server. "
    "GitHub updates must first reach the deployed app."
)

st.caption(
    "Averages use available numeric observations. "
    "Missing values are excluded and zeros are included. "
    "Downloads are optional."
)


if section == "NG Swing GDD master output":
    st.header("NG Swing GDD master output")

    master_files = sorted(
        (ROOT / "master").rglob("*.csv")
    )

    if not master_files:
        st.error("No CSV files were found in the master folder.")
        st.stop()

    preferred_files = [
        p for p in master_files
        if p.name.lower().startswith("master_ng_swing_gdd")
    ]

    default_file = (
        preferred_files[0]
        if preferred_files
        else master_files[0]
    )

    master_path = st.sidebar.selectbox(
        "Master file",
        options=master_files,
        index=master_files.index(default_file),
        format_func=lambda p: str(p.relative_to(ROOT)),
    )

    try:
        with st.spinner("Reading master output..."):
            data = load_master(
                str(master_path),
                master_path.stat().st_mtime_ns,
                master_path.stat().st_size,
            )

    except Exception as error:
        st.error(str(error))
        st.stop()

    st.caption(
        f"{len(data):,} dates | "
        f"{len(data.columns):,} hub series | "
        f"{data.index.min():%b %d, %Y} through "
        f"{data.index.max():%b %d, %Y}"
    )

    st.info(
        "These are values from the daily process's master output. "
        "The file does not identify actual versus modeled values "
        "or provide a historical run/as-of date."
    )

    start, end = select_dates(
        data.index,
        "master",
    )

    hubs = data.columns.tolist()

    defaults = [
        hub for hub in [
            "HSC",
            "Henry",
            "Waha",
            "TETCO-M3",
        ]
        if hub in hubs
    ]

    selected_hubs = st.sidebar.multiselect(
        "Master-output hubs",
        hubs,
        default=defaults or hubs[:1],
    )

    if not selected_hubs:
        st.info("Select at least one hub.")
        st.stop()

    selected = data.loc[
        (data.index >= start)
        & (data.index <= end),
        selected_hubs,
    ]

    st.caption(
        "Each hub has its own average row. "
        "Different hubs are not blended into one price."
    )

    price_report(
        selected,
        "master_report",
        selected_strip_average=False,
    )


else:
    st.header("ICE settlements")

    code_text = st.sidebar.text_input(
        "ICE codes to load",
        value="",
        help=(
            "Blank loads all instruments. "
            "Optionally enter comma-separated codes."
        ),
    )

    codes = tuple(
        sorted({
            x.strip().upper()
            for x in code_text.split(",")
            if x.strip()
        })
    )

    paths = [
        p for p in (ROOT / "input").rglob("icecleared*.xlsx")
        if not p.name.startswith("~$")
    ]

    if not paths:
        st.error("No icecleared*.xlsx files found in input.")
        st.stop()

    try:
        with st.spinner(
            "Loading settlement history. "
            "The initial full import may take several minutes..."
        ):
            data = load_ice(
                manifest(paths),
                codes,
            )

    except Exception as error:
        st.error(str(error))
        st.stop()

    st.caption(
        f"{len(paths):,} files | "
        f"{len(data):,} unique records | "
        f"{data['contract'].nunique():,} ICE codes | "
        f"{data['trade_date'].min():%b %d, %Y} through "
        f"{data['trade_date'].max():%b %d, %Y}"
    )

    start, end = select_dates(
        data["trade_date"],
        "ice",
    )

    available_codes = sorted(
        data["contract"].unique().tolist()
    )

    default_code = next(
        (
            c for c in ["HXS", "FQT", "TPH"]
            if c in available_codes
        ),
        available_codes[0],
    )

    view = multiselect_filter(
        data,
        "contract",
        "ICE code",
        "ice_code",
        default=[default_code],
    )

    view = multiselect_filter(
        view,
        "product",
        "Product",
        "ice_product",
    )

    view = multiselect_filter(
        view,
        "hub",
        "Hub",
        "ice_hub",
    )

    view = multiselect_filter(
        view,
        "contract_type",
        "Contract type",
        "ice_type",
    )

    strips = sorted(
        view["strip"].dropna().unique().tolist()
    )

    selected_strips = st.sidebar.multiselect(
        "Delivery strips",
        strips,
        default=strips[:1],
        format_func=lambda x: (
            pd.Timestamp(x).strftime("%b %d, %Y")
        ),
        help="Clear the selection to include all available strips.",
    )

    if selected_strips:
        view = view[
            view["strip"].isin(selected_strips)
        ]

    strikes = sorted(
        view["strike"].dropna().unique().tolist()
    )

    selected_strikes = st.sidebar.multiselect(
        "Strike, if applicable",
        strikes,
        help="Empty includes all strikes and records with no strike.",
    )

    if selected_strikes:
        view = view[
            view["strike"].isin(selected_strikes)
        ]

    metric_label = st.sidebar.selectbox(
        "Price field",
        ["Settlement price", "Net change"],
    )

    metric = (
        "settlement_price"
        if metric_label == "Settlement price"
        else "net_change"
    )

    view = view[
        view["trade_date"].between(start, end)
    ].copy()

    if view.empty:
        st.info("No records match the filters.")
        st.stop()

    # Expirations can differ by delivery strip.
    # Keep expiration in the table-column identity rather
    # than splitting the strip comparison into separate reports.
    family_columns = [
        c for c in ICE_FAMILY
        if c != "expiration_date"
    ]

    grouped = view.groupby(
        family_columns,
        observed=True,
        dropna=False,
        sort=True,
    )

    if grouped.ngroups > 20:
        st.info(
            "More than 20 contract groups match. "
            "Narrow the code, product, hub, type, or strike filters "
            "to display price tables."
        )

    else:
        for position, (labels, group) in enumerate(grouped):
            if not isinstance(labels, tuple):
                labels = (labels,)

            heading = " | ".join(
                f"{column}: {label}"
                for column, label in zip(
                    family_columns,
                    labels,
                )
                if pd.notna(label) and str(label) != ""
            )

            st.subheader(heading)

            group = group.copy()

            # Usually one expiration per strip. If the source
            # contains several, expose them as separate columns.
            expirations_per_strip = (
                group.groupby(
                    "strip",
                    observed=True,
                    dropna=False,
                )["expiration_date"]
                .nunique(dropna=False)
            )

            multiple_expirations = (
                expirations_per_strip > 1
            ).any()

            group["delivery_column"] = (
                group["strip"].dt.strftime("%Y-%m-%d")
            )

            if multiple_expirations:
                group["delivery_column"] = (
                    group["delivery_column"]
                    + " | expires "
                    + group["expiration_date"]
                    .dt.strftime("%Y-%m-%d")
                    .fillna("unspecified")
                )

            if group.duplicated(
                ["trade_date", "delivery_column"]
            ).any():
                st.error(
                    "Multiple instrument identities map to one "
                    "table cell. Narrow the filters."
                )
                continue

            wide = group.pivot(
                index="trade_date",
                columns="delivery_column",
                values=metric,
            )

            wide.columns.name = None

            st.caption(
                f"Price field: {metric_label}. "
                "The strip average equally weights the selected "
                "delivery columns. It is not volume-weighted "
                "or weighted by days in the delivery month."
            )

            price_report(
                wide,
                f"ice_group_{position}",
                selected_strip_average=True,
            )

    with st.expander("Underlying ICE records and source files"):
        st.dataframe(
            view,
            use_container_width=True,
            hide_index=True,
        )

        st.download_button(
            "Download underlying ICE records",
            view.to_csv(index=False).encode("utf-8-sig"),
            "ice_underlying_records.csv",
            "text/csv",
        )
