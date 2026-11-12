from pathlib import Path
import hashlib
import json

import duckdb
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
INPUT = ROOT / "input"
ARCHIVE = ROOT / "ice_archive"
PARTS = ARCHIVE / "parts"

PARTS.mkdir(parents=True, exist_ok=True)

MANIFEST = ARCHIVE / "manifest.json"
VERSION = 1

TEXT = [
    "contract",
    "product",
    "hub",
    "contract_type",
    "product_id",
]

DATES = [
    "trade_date",
    "strip",
    "expiration_date",
]

NUMBERS = [
    "strike",
    "settlement_price",
    "net_change",
]

IDENTITY = [
    "trade_date",
    "contract",
    "product",
    "hub",
    "strip",
    "contract_type",
    "strike",
    "expiration_date",
    "product_id",
]

REQUIRED = TEXT + DATES + NUMBERS

old = (
    json.loads(MANIFEST.read_text())
    if MANIFEST.exists()
    else {}
)

if old.get("version") != VERSION:
    old = {}

previous = old.get("files", {})
current = {}

files = sorted(
    path
    for path in INPUT.rglob("icecleared*.xlsx")
    if not path.name.startswith("~$")
)

if not files:
    raise RuntimeError("No icecleared*.xlsx files found in input.")

converted = 0

for path in files:
    relative = path.relative_to(ROOT).as_posix()

    digest = hashlib.sha256(path.read_bytes()).hexdigest()

    # Stable output name: replacing a source replaces its Parquet.
    part_name = hashlib.sha256(
        relative.encode()
    ).hexdigest()[:24] + ".parquet"

    destination = PARTS / part_name

    info = {
        "sha256": digest,
        "part": part_name,
    }

    current[relative] = info

    if (
        previous.get(relative) == info
        and destination.exists()
    ):
        print(f"Unchanged: {relative}", flush=True)
        continue

    print(f"Converting: {relative}", flush=True)

    df = pd.read_excel(
        path,
        engine="openpyxl",
        dtype=object,
    )

    df.columns = [
        str(c).strip().lower().replace(" ", "_")
        for c in df.columns
    ]

    missing = set(REQUIRED) - set(df.columns)

    if missing:
        raise RuntimeError(
            f"{relative}: missing columns {sorted(missing)}"
        )

    df = df[REQUIRED].dropna(how="all").copy()

    for column in TEXT:
        df[column] = (
            df[column].fillna("").astype(str).str.strip()
        )

    df["contract"] = df["contract"].str.upper()

    for column in DATES + NUMBERS:
        original = df[column]

        present = (
            original.notna()
            & original.astype(str).str.strip().ne("")
        )

        if column in DATES:
            parsed = pd.to_datetime(
                original.where(present),
                errors="coerce",
            )
        else:
            parsed = pd.to_numeric(
                original.where(present),
                errors="coerce",
            ).astype("float64")

        if (present & parsed.isna()).any():
            raise RuntimeError(
                f"{relative}: invalid values in {column}"
            )

        df[column] = parsed

    if (
        df["trade_date"].isna()
        | df["strip"].isna()
        | df["contract"].eq("")
    ).any():
        raise RuntimeError(
            f"{relative}: missing trade date, strip, or ICE code"
        )

    df["source_file"] = relative

    df = df.sort_values(
        ["contract", "hub", "strip", "trade_date"]
    )

    df.to_parquet(
        destination,
        index=False,
        compression="zstd",
    )

    converted += 1

# Remove converted parts whose source files were deleted.
expected = {item["part"] for item in current.values()}

for path in PARTS.glob("*.parquet"):
    if path.name not in expected:
        path.unlink()

con = duckdb.connect()
con.execute("SET memory_limit='2GB'")

parquet_glob = (PARTS / "*.parquet").as_posix()

con.read_parquet(
    parquet_glob,
    union_by_name=True,
).create_view("raw")

key_sql = ", ".join(IDENTITY)

# NULL counts as a distinct value for conflict detection.
conflicts = con.execute(
    f"""
    SELECT {key_sql}
    FROM raw
    GROUP BY {key_sql}
    HAVING
      COUNT(DISTINCT settlement_price)
        + CASE WHEN COUNT(*) > COUNT(settlement_price)
               THEN 1 ELSE 0 END > 1
      OR
      COUNT(DISTINCT net_change)
        + CASE WHEN COUNT(*) > COUNT(net_change)
               THEN 1 ELSE 0 END > 1
    LIMIT 10
    """
).df()

if not conflicts.empty:
    raise RuntimeError(
        "Conflicting duplicate contract/date records:\n"
        + conflicts.to_string(index=False)
    )

# Catalog contains instrument identities and date coverage,
# not the full price history.
catalog = con.execute(
    """
    SELECT
      contract, product, hub, strip, contract_type,
      strike, expiration_date, product_id,
      MIN(trade_date) AS first_trade_date,
      MAX(trade_date) AS last_trade_date
    FROM raw
    GROUP BY
      contract, product, hub, strip, contract_type,
      strike, expiration_date, product_id
    ORDER BY contract, hub, strip
    """
).df()

catalog.to_parquet(
    ARCHIVE / "catalog.parquet",
    index=False,
    compression="zstd",
)

summary = {
    "source_files": len(files),
    "converted_this_run": converted,
    "instrument_count": len(catalog),
    "first_trade_date": str(catalog["first_trade_date"].min()),
    "last_trade_date": str(catalog["last_trade_date"].max()),
}

(ARCHIVE / "summary.json").write_text(
    json.dumps(summary, indent=2)
)

MANIFEST.write_text(
    json.dumps(
        {
            "version": VERSION,
            "files": current,
        },
        indent=2,
        sort_keys=True,
    )
)

print(json.dumps(summary, indent=2))
con.close()
