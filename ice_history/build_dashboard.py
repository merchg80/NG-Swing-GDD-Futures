from pathlib import Path
import json
import os
import sqlite3

import pandas as pd
from plotly.offline import get_plotlyjs


INPUT = Path("input")
OUTPUT = Path("ice_history_output")
OUTPUT.mkdir(exist_ok=True)

# Blank means import every ICE code.
requested = {
    x.strip().upper()
    for x in os.getenv("ICE_CODES", "").split(",")
    if x.strip()
}

required = [
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

date_columns = ["trade_date", "strip", "expiration_date"]
numeric_columns = ["strike", "settlement_price", "net_change"]
text_columns = ["hub", "product", "contract", "contract_type", "product_id"]

frames = []
files = sorted(INPUT.rglob("icecleared*.xlsx"))

if not files:
    raise RuntimeError("No icecleared*.xlsx files were found in input.")

for path in files:
    if path.name.startswith("~$"):
        continue

    print(f"Reading {path}", flush=True)
    df = pd.read_excel(path, engine="openpyxl", dtype=object)
    df.columns = [
        str(c).strip().lower().replace(" ", "_")
        for c in df.columns
    ]

    missing = set(required) - set(df.columns)
    if missing:
        raise RuntimeError(
            f"{path}: missing columns {sorted(missing)}"
        )

    df = df[required].copy()
    df = df.dropna(how="all")

    for column in text_columns:
        df[column] = df[column].fillna("").astype(str).str.strip()

    df["contract"] = df["contract"].str.upper()

    if requested:
        df = df[df["contract"].isin(requested)].copy()

    if df.empty:
        continue

    for column in date_columns:
        original = df[column]
        parsed = pd.to_datetime(original, errors="coerce")
        invalid = original.notna() & parsed.isna()

        if invalid.any():
            raise RuntimeError(
                f"{path}: invalid {column}: "
                f"{original[invalid].head(5).tolist()}"
            )

        df[column] = parsed.dt.strftime("%Y-%m-%d").fillna("")

    for column in numeric_columns:
        original = df[column]
        parsed = pd.to_numeric(original, errors="coerce")
        invalid = (
            original.notna()
            & original.astype(str).str.strip().ne("")
            & parsed.isna()
        )

        if invalid.any():
            raise RuntimeError(f"{path}: invalid number in {column}")

        df[column] = parsed

    bad_identity = (
        df["trade_date"].eq("")
        | df["contract"].eq("")
        | df["strip"].eq("")
    )
    if bad_identity.any():
        raise RuntimeError(
            f"{path}: {int(bad_identity.sum())} rows lack "
            "trade date, contract, or strip."
        )

    df["source_file"] = path.as_posix()
    frames.append(df)

if not frames:
    raise RuntimeError(
        "No matching pricing rows. Check the ICE codes entered."
    )

data = pd.concat(frames, ignore_index=True)

identity = [
    "trade_date",
    "hub",
    "product",
    "strip",
    "contract",
    "contract_type",
    "strike",
    "expiration_date",
    "product_id",
]

# Check duplicate contract/date records before removing duplicates.
counts = (
    data.groupby(identity, dropna=False)[
        ["settlement_price", "net_change"]
    ]
    .nunique(dropna=False)
)

conflicts = counts[(counts > 1).any(axis=1)]

if not conflicts.empty:
    conflicts.reset_index().to_csv(
        OUTPUT / "conflicting_keys.csv", index=False
    )
    raise RuntimeError(
        f"{len(conflicts)} conflicting contract/date records. "
        "See conflicting_keys.csv in the diagnostic artifact."
    )

# Retain every source-file reference, but count each price only once.
sources = (
    data.groupby(identity, dropna=False)["source_file"]
    .agg(lambda values: "; ".join(sorted(set(values))))
    .reset_index()
)

data = data.drop(columns="source_file").drop_duplicates(
    subset=identity
)
data = data.merge(
    sources, on=identity, how="left", validate="one_to_one"
)
data = data.sort_values(
    ["contract", "hub", "strip", "trade_date"]
).reset_index(drop=True)

db_path = OUTPUT / "ice_history.sqlite"

if db_path.exists():
    db_path.unlink()

with sqlite3.connect(db_path) as connection:
    data.to_sql(
        "settlements", connection, index=False, if_exists="replace"
    )
    connection.execute(
        "CREATE INDEX price_lookup ON settlements "
        "(contract, hub, strip, trade_date)"
    )

metadata = {
    "files_read": len(frames),
    "unique_rows": len(data),
    "first_trade_date": data["trade_date"].min(),
    "last_trade_date": data["trade_date"].max(),
    "codes": sorted(data["contract"].unique().tolist()),
}

(OUTPUT / "import_summary.json").write_text(
    json.dumps(metadata, indent=2), encoding="utf-8"
)

# Column-oriented encoding avoids repeating field names on every row.
payload = json.dumps(
    {
        "columns": data.columns.tolist(),
        "rows": json.loads(data.to_json(orient="values")),
        "metadata": metadata,
    },
    separators=(",", ":"),
).replace("<", "\\u003c")

HTML = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>ICE Settlement History</title>
<style>
body{font:15px Arial,sans-serif;margin:28px;color:#182d40;background:#f4f7fa}
h1{margin-bottom:8px}h2{margin-top:26px}
.note{color:#526678;line-height:1.5}
.filters{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:14px;
background:white;padding:20px;border-radius:10px;margin:20px 0}
label{display:block;font-weight:bold;font-size:13px}
select,input{box-sizing:border-box;width:100%;margin-top:6px;padding:9px;
border:1px solid #bac8d4;border-radius:5px;background:white}
button{padding:10px 16px;background:#124e78;color:white;border:0;
border-radius:5px;cursor:pointer;margin:6px 8px 6px 0}
.panel{background:white;padding:16px;border-radius:10px;margin:16px 0}
.scroll{overflow:auto;max-height:500px}
table{border-collapse:collapse;width:100%;white-space:nowrap;font-size:13px}
th,td{padding:9px 12px;border-bottom:1px solid #e0e6ec;text-align:right}
th{background:#eaf0f5;position:sticky;top:0}
th:first-child,td:first-child{text-align:left}
#status{font-weight:bold}
</style>
<script>__PLOTLY__</script>
</head>
<body>
<h1>ICE Settlement History</h1>
<div class="note" id="coverage"></div>
<p class="note">
Strip means delivery period. Dates below mean trading dates.
Averages are arithmetic averages of available observations, not volume-weighted.
Each distinct contract, hub, product, type, strike, expiration and product ID
is calculated separately. Missing prices are excluded; zero prices are included.
</p>
<div class="filters">
<label>ICE code<select id="contract"></select></label>
<label>Product<select id="product"></select></label>
<label>Hub<select id="hub"></select></label>
<label>Contract type<select id="contract_type"></select></label>
<label>Delivery strip<select id="strip"></select></label>
<label>Price field<select id="metric">
<option value="settlement_price">Settlement price</option>
<option value="net_change">Net change</option>
</select></label>
<label>First trade date<input type="date" id="start"></label>
<label>Last trade date<input type="date" id="end"></label>
</div>
<button id="apply">Update report</button>
<button id="dailyDownload">Download prices CSV</button>
<button id="monthlyDownload">Download monthly averages CSV</button>
<button id="periodDownload">Download period averages CSV</button>
<p id="status"></p>
<div class="panel"><div id="chart"></div></div>
<h2>Selected-period averages</h2>
<div class="panel scroll" id="period"></div>
<h2>Monthly averages by trade month</h2>
<div class="panel scroll" id="monthly"></div>
<h2>Underlying prices</h2>
<p class="note">The screen shows the first 1,000 matching rows.
The prices CSV includes every matching row and its source files.</p>
<div class="panel scroll" id="daily"></div>
<script id="dataset" type="application/json">__DATA__</script>
<script>
"use strict";
const packet=JSON.parse(document.getElementById("dataset").textContent);
const columns=packet.columns;
const index=Object.fromEntries(columns.map((name,i)=>[name,i]));
const rows=packet.rows;
const value=(row,name)=>row[index[name]];
const element=id=>document.getElementById(id);
const filters=["contract","product","hub","contract_type","strip"];
const identity=["contract","product","hub","strip","contract_type",
                "strike","expiration_date","product_id"];
const number=x=>typeof x==="number"&&Number.isFinite(x);
const fmt=x=>number(x)?x.toFixed(4):(x??"");
let selected=[],monthly=[],period=[];

element("coverage").textContent=
`${packet.metadata.files_read} source files | ${rows.length.toLocaleString()} unique rows | `+
`${packet.metadata.first_trade_date} through ${packet.metadata.last_trade_date}`;

element("start").value=packet.metadata.first_trade_date;
element("end").value=packet.metadata.last_trade_date;

function refill(position){
  for(let k=position;k<filters.length;k++){
    const name=filters[k], control=element(name), old=control.value;
    const candidates=rows.filter(row=>
      filters.slice(0,k).every(f=>
        !element(f).value||String(value(row,f)??"")===element(f).value));
    const options=[...new Set(candidates.map(row=>String(value(row,name)??"")))]
      .filter(Boolean).sort();
    control.replaceChildren(new Option("All",""),
      ...options.map(x=>new Option(x,x)));
    if(options.includes(old))control.value=old;
  }
}

refill(0);
for(const name of filters){
  element(name).addEventListener("change",()=>{
    refill(filters.indexOf(name)+1);
  });
}
const preferred=["HXS","FQT","TPH","FQH","UCS"];
const first=preferred.find(code=>rows.some(row=>value(row,"contract")===code));
element("contract").value=first||String(value(rows[0],"contract"));
refill(1);
if([...element("contract_type").options].some(o=>o.value==="F")){
  element("contract_type").value="F";refill(4);
}
const strips=[...element("strip").options].map(o=>o.value).filter(Boolean);
if(strips.length)element("strip").value=strips[0];

function table(target,headers,records){
  const t=document.createElement("table");
  const head=t.createTHead().insertRow();
  headers.forEach(h=>{const cell=document.createElement("th");
    cell.textContent=h;head.appendChild(cell);});
  const body=t.createTBody();
  records.forEach(record=>{const tr=body.insertRow();
    record.forEach(x=>{tr.insertCell().textContent=fmt(x);});});
  element(target).replaceChildren(t);
}

function aggregate(input,metric,byMonth){
  const groups=new Map();
  for(const row of input){
    const price=value(row,metric);
    const labels=identity.map(f=>value(row,f));
    if(byMonth)labels.push(value(row,"trade_date").slice(0,7));
    const key=JSON.stringify(labels);
    if(!groups.has(key))groups.set(key,{labels,values:[]});
    if(number(price))groups.get(key).values.push(price);
  }
  return [...groups.values()].map(g=>{
    const v=g.values;
    return [...g.labels,
      v.length?v.reduce((a,b)=>a+b,0)/v.length:null,
      v.length,
      v.length?Math.min(...v):null,
      v.length?Math.max(...v):null];
  }).sort((a,b)=>JSON.stringify(a).localeCompare(JSON.stringify(b)));
}

function update(){
  const start=element("start").value,end=element("end").value;
  if(!start||!end||start>end){
    element("status").textContent="Enter a valid start and end date.";return;
  }
  selected=rows.filter(row=>
    value(row,"trade_date")>=start&&value(row,"trade_date")<=end&&
    filters.every(f=>!element(f).value||
      String(value(row,f)??"")===element(f).value));
  const metric=element("metric").value;
  period=aggregate(selected,metric,false);
  monthly=aggregate(selected,metric,true);
  table("period",[...identity,"average","observations","minimum","maximum"],period);
  table("monthly",[...identity,"trade_month","average","observations","minimum","maximum"],monthly);
  table("daily",columns,selected.slice(0,1000));

  const groups=new Map();
  selected.forEach(row=>{
    const labels=identity.map(f=>value(row,f));
    const key=JSON.stringify(labels);
    if(!groups.has(key))groups.set(key,{labels,items:[]});
    if(number(value(row,metric)))groups.get(key).items.push(row);
  });
  const traces=[];
  let count=0;
  for(const group of groups.values()){
    if(count++>=30)break;
    group.items.sort((a,b)=>value(a,"trade_date").localeCompare(value(b,"trade_date")));
    const label=group.labels.filter(x=>x!==null&&x!=="").join(" | ");
    traces.push({
      x:group.items.map(row=>value(row,"trade_date")),
      y:group.items.map(row=>value(row,metric)),
      name:label,type:"scatter",mode:"lines+markers",marker:{size:4}
    });
    const key=JSON.stringify(group.labels);
    const average=period.find(p=>JSON.stringify(p.slice(0,identity.length))===key);
    if(average&&number(average[identity.length])){
      traces.push({
        x:[start,end],
        y:[average[identity.length],average[identity.length]],
        name:label+" average",type:"scatter",mode:"lines",
        line:{dash:"dot",width:1},showlegend:false
      });
    }
  }
  Plotly.react("chart",traces,{
    title:metric==="settlement_price"?"Settlement history":"Net change history",
    height:520,margin:{t:55,r:25,b:80,l:75},
    xaxis:{title:"Trade date",type:"date"},
    yaxis:{title:metric.replaceAll("_"," ")},
    legend:{orientation:"h",y:-0.25},hovermode:"closest"
  },{responsive:true,displaylogo:false});
  const priced=selected.filter(r=>number(value(r,metric))).length;
  element("status").textContent=
    `${selected.length.toLocaleString()} matching rows; `+
    `${priced.toLocaleString()} numeric observations; ${period.length} separate series.`+
    (groups.size>30?" Chart shows the first 30 series. Narrow the filters to chart the others.":"");
}

function download(filename,headers,records){
  const escape=x=>{
    let s=String(x??"");
    if(typeof x==="string"&&/^[=+@\t\r-]/.test(s))s="'"+s;
    return '"'+s.replaceAll('"','""')+'"';
  };
  const csv="\ufeff"+[headers,...records]
    .map(row=>row.map(escape).join(",")).join("\r\n");
  const url=URL.createObjectURL(new Blob([csv],{type:"text/csv;charset=utf-8"}));
  const a=document.createElement("a");a.href=url;a.download=filename;a.click();
  setTimeout(()=>URL.revokeObjectURL(url),1000);
}
element("apply").onclick=update;
element("dailyDownload").onclick=()=>download("ice_prices.csv",columns,selected);
element("monthlyDownload").onclick=()=>download("ice_monthly_averages.csv",
  [...identity,"trade_month","average","observations","minimum","maximum"],monthly);
element("periodDownload").onclick=()=>download("ice_period_averages.csv",
  [...identity,"average","observations","minimum","maximum"],period);
update();
</script>
</body>
</html>
"""

html = HTML.replace("__PLOTLY__", get_plotlyjs()).replace(
    "__DATA__", payload
)

dashboard_path = OUTPUT / "ICE_History_Dashboard.html"
dashboard_path.write_text(html, encoding="utf-8")

print(json.dumps(metadata, indent=2))
print(
    "Dashboard size:",
    round(dashboard_path.stat().st_size / 1024 / 1024, 1),
    "MB",
)
print("Created dashboard and SQLite database.")
