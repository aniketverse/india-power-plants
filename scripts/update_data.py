#!/usr/bin/env python3
"""
Monitor CEA / National Power Portal published reports and append newly detected
utility generating stations to data/cea_additions.csv.

Safety design:
- Does not delete prior records when a source report temporarily omits a station.
- Excludes industrial sites using both station and owner/utility names.
- Automatically accepts Central/State sector records.
- Holds ambiguous private-sector records out of the public inventory unless
  the report explicitly identifies them as utility/IPP or owner is allow-listed.
- If source structure changes or parsing yields no usable rows, exits without
  replacing the existing feed.
"""
from __future__ import annotations
import csv, io, re, sys, hashlib, json
from pathlib import Path
from datetime import datetime, timezone
from urllib.parse import urljoin
import requests
import pandas as pd
from bs4 import BeautifulSoup

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
CSV_PATH = DATA / "cea_additions.csv"
LOG_PATH = DATA / "update_log.json"
REPORTS_URL = "https://npp.gov.in/publishedReports"
HEADERS = {"User-Agent": "Aniketverse-PowerStations-Monitor/1.0 (public research dashboard)"}
TIMEOUT = 45

INDUSTRY = re.compile(
    r"\b(cement|clinker|steel|sponge iron|iron works|iron ore|pig iron|ferro ?alloy|"
    r"rolling mill|foundry|smelter|mine|mining|colliery|coal washery|mineral|bauxite|"
    r"limestone|quarry|refinery|petroleum|petrochemical|oil terminal|gas processing|"
    r"lng terminal|lpg bottling|chemical|fertili[sz]er|ammonia|urea|chlor.?alkali|"
    r"caustic soda|pharma|pharmaceutical|pesticide|paper mill|paper plant|pulp|kraft|"
    r"board mill|packaging mill|sugar mill|sugar factory|distillery|ethanol plant|"
    r"molasses|textile|spinning mill|weaving|fabric|garment|industrial park|"
    r"industrial estate|manufacturing|factory|aluminium|aluminum|copper|zinc|lead|"
    r"metal plant|metal works|port|dock|airport|railway workshop|locomotive works|"
    r"tyre|tire|rubber|automobile|auto plant|engineering works|glass works|ceramic|"
    r"refractory|food processing|brewery|ultratech|acc limited|ambuja cement|shree cement|dalmia cement|grasim|hindalco|tata steel|jsw steel|jindal stainless|vedanta|hindustan zinc|balco|nalco|indian oil|bharat petroleum|bpcl|hpcl|ongc|gail|reliance industries|aditya birla|coromandel|chambal fertilisers|national fertilizers|paradeep phosphates|itc limited|jk paper|west coast paper|tnpl|trident|bajaj hindusthan|shree renuka|eid parry)\b", re.I
)
UTILITY_OWNERS = [
    "ntpc", "nhpc", "npcil", "sjvn", "thdc", "neepco", "dvc", "bbmb", "nlc",
    "maharashtra state power generation", "mahagenco", "gujarat state electricity",
    "gujarat state electricity corporation", "rajasthan rajya vidyut", "rvunl",
    "madhya pradesh power generating", "mppgcl", "uttar pradesh rajya vidyut",
    "up rvunl", "haryana power generation", "hpgcl", "punjab state power",
    "pspcl", "karnataka power corporation", "kpcl", "telangana state power",
    "tsspdcl", "tsgen", "andhra pradesh power generation", "apgenco",
    "tata power", "adani power", "torrent power", "cesc", "jsw energy",
    "jindal power", "rattanindia", "gmr energy", "reliance power", "lanco",
    "greenko", "renew", "azure power", "acme solar", "sembcorp", "hero future",
    "waaree", "ntpc green"
]
FUEL_WORDS = [
    (re.compile(r"ther\s*\(gt\)|ccgt|ccpp|gas|combined cycle", re.I), "Gas"),
    (re.compile(r"ther\s*\(dg\)|diesel|oil", re.I), "Oil"),
    (re.compile(r"coal|lignite|thermal", re.I), "Coal"),
    (re.compile(r"hydro|hydel|pumped", re.I), "Hydro"),
    (re.compile(r"nuclear|atomic", re.I), "Nuclear"),
    (re.compile(r"solar", re.I), "Solar"),
    (re.compile(r"wind", re.I), "Wind"),
    (re.compile(r"biomass|bagasse", re.I), "Biomass"),
    (re.compile(r"waste", re.I), "Waste"),
]

def clean(v):
    if pd.isna(v): return ""
    return re.sub(r"\s+", " ", str(v)).strip()

def key(v):
    return re.sub(r"[^a-z0-9]", "", clean(v).lower())

def fuel_group(v):
    for pat, name in FUEL_WORDS:
        if pat.search(clean(v)): return name
    return "Other"

def industrial(name, owner):
    return bool(INDUSTRY.search(f"{clean(name)} {clean(owner)}"))

def fetch_report_links(html, base_url):
    soup = BeautifulSoup(html, "html.parser")

    targets = {
        "daily_stationwise": re.compile(
            r"02\s*Region\s*Wise.*?Station\s*Wise.*?Unit\s*Wise\s*Generation\s*Report",
            re.I | re.S
        ),
        "monthly_stationwise": re.compile(
            r"16\s*Region\s*wise\s*State\s*wise\s*and\s*Station\s*wise\s*Detail",
            re.I | re.S
        ),
        "installed_capacity": re.compile(
            r"01\s*All\s*India\s*Installed\s*Capacity\s*of\s*Power\s*Stations",
            re.I | re.S
        ),
    }

    found = {}

    for anchor in soup.find_all("a", href=True):
        link_text = anchor.get_text(" ", strip=True)
        href = anchor["href"]

        is_excel = (
            re.search(r"\b(xls|excel|xlsx)\b", link_text, re.I)
            or re.search(r"\.(xlsx?|xlsm)(?:\?|$)", href, re.I)
        )

        if not is_excel:
            continue

        parent = anchor

        for _ in range(12):
            if parent is None:
                break

            context = parent.get_text(" ", strip=True)

            matches = [
                label for label, pattern in targets.items()
                if pattern.search(context)
            ]

            if len(matches) == 1:
                label = matches[0]
                found.setdefault(label, urljoin(base_url, href))
                break

            if len(matches) > 1:
                break

            parent = parent.parent

    return found
  
def find_col(headers, patterns):
    for i, h in enumerate(headers):
        if any(re.search(p, clean(h), re.I) for p in patterns): return i
    return None


CEA_STATE_MAP = {
    "ANDHRA PRADESH": "Andhra Pradesh",
    "ARUNACHAL PRADESH": "Arunachal Pradesh",
    "ASSAM": "Assam",
    "BIHAR": "Bihar",
    "CHHATISGARH": "Chhattisgarh",
    "CHHATTISGARH": "Chhattisgarh",
    "GOA": "Goa",
    "GUJARAT": "Gujarat",
    "HARYANA": "Haryana",
    "HIMACHAL PRADESH": "Himachal Pradesh",
    "JHARKHAND": "Jharkhand",
    "KARNATAKA": "Karnataka",
    "KERALA": "Kerala",
    "MADHYA PRADESH": "Madhya Pradesh",
    "MAHARASHTRA": "Maharashtra",
    "MANIPUR": "Manipur",
    "MEGHALAYA": "Meghalaya",
    "MIZORAM": "Mizoram",
    "NAGALAND": "Nagaland",
    "ODISHA": "Odisha",
    "ORISSA": "Odisha",
    "PUNJAB": "Punjab",
    "RAJASTHAN": "Rajasthan",
    "SIKKIM": "Sikkim",
    "TAMIL NADU": "Tamil Nadu",
    "TELANGANA": "Telangana",
    "TRIPURA": "Tripura",
    "UTTAR PRADESH": "Uttar Pradesh",
    "UTTARAKHAND": "Uttarakhand",
    "WEST BENGAL": "West Bengal",
    "JAMMU AND KASHMIR": "Jammu and Kashmir",
    "LADAKH": "Ladakh",
    "DELHI": "Delhi",
}

CEA_REGION_HEADINGS = {
    "NORTHERN", "NORTHERN REGION",
    "WESTERN", "WESTERN REGION",
    "SOUTHERN", "SOUTHERN REGION",
    "EASTERN", "EASTERN REGION",
    "NORTH EASTERN", "NORTH-EASTERN",
    "NORTH EAST REGION", "BHUTAN IMP.", "BHUTAN"
}


def parse_cea_layout(df, source_label):
    """Parse CEA JasperReports XLS with State/Sector/Type group rows."""

    preview = " ".join(
        clean(x) for x in df.head(5).to_numpy().ravel()
    )

    if "GENERATION REPORT" not in preview.upper():
        return []

    date_match = re.search(
        r"REPORT\s+FOR\s+(\d{1,2}/\d{1,2}/\d{4})",
        preview, re.I
    )

    source_as_of = ""
    if date_match:
        try:
            source_as_of = datetime.strptime(
                date_match.group(1), "%d/%m/%Y"
            ).date().isoformat()
        except ValueError:
            pass

    current_state = ""
    current_sector = ""
    current_fuel = ""
    records = []

    for row in df.itertuples(index=False, name=None):
        vals = [clean(x) for x in row]

        if not vals:
            continue

        name = vals[0]
        upper = name.upper().strip()

        if not upper:
            continue

        if upper in CEA_REGION_HEADINGS:
            current_state = ""
            current_sector = ""
            current_fuel = ""
            continue

        if upper in CEA_STATE_MAP:
            current_state = CEA_STATE_MAP[upper]
            current_sector = ""
            current_fuel = ""
            continue

        if "SECTOR:" in upper:
            current_sector = vals[4] if len(vals) > 4 else ""
            current_fuel = ""
            continue

        if "TYPE:" in upper:
            current_fuel = vals[4] if len(vals) > 4 else ""
            continue

        if not current_state or not name:
            continue

        if upper == "UNIT" or "TOTAL" in upper:
            continue

        # Column H (index 7) contains monitored capacity in MW.
        if len(vals) <= 7 or not vals[7]:
            continue

        try:
            capacity = float(vals[7].replace(",", ""))
        except (TypeError, ValueError):
            continue

        if industrial(name, ""):
            continue

        sector_text = current_sector.lower()

        is_central_state = bool(re.search(
            r"\b(central sector|state sector|state government)\b",
            sector_text
        ))

        is_utility_ipp = bool(re.search(
            r"\b(utility|ipp|independent power producer)\b",
            sector_text
        ))

        # Some private-sector plant names contain the utility owner.
        known_owner = any(
            owner in name.lower() for owner in UTILITY_OWNERS
        )

        if not (is_central_state or is_utility_ipp or known_owner):
            continue

        records.append({
            "name": name,
            "state": current_state,
            "rawFuel": current_fuel or "Other",
            "capacity": f"{capacity:g}",
            "owner": "",
            "year": "",
            "lat": "",
            "lon": "",
            "source": f"CEA/NPP {source_label}",
            "source_as_of": source_as_of
        })

    return records

def parse_workbook(content, source_label):
    # pandas selects xlrd for legacy .xls and openpyxl for .xlsx.
    book = pd.read_excel(io.BytesIO(content), sheet_name=None, header=None, dtype=str)
    records = []
    for sheet_name, df in book.items():
        if df.empty: continue
        df = df.fillna("")
        grid_records = parse_cea_layout(df, source_label)
        if grid_records:
            records.extend(grid_records)
            continue
        header_idx = None
        cols = None
        for ridx in range(min(len(df), 45)):
            cells = [clean(x) for x in df.iloc[ridx].tolist()]
            station_col = find_col(cells, [r"\bstation\b", r"power station", r"project name", r"plant name"])
            state_col = find_col(cells, [r"\bstate\b"])
            if station_col is not None and state_col is not None:
                header_idx, cols = ridx, cells
                break
        if header_idx is None: continue
        station_col = find_col(cols, [r"\bstation\b", r"power station", r"project name", r"plant name"])
        state_col = find_col(cols, [r"\bstate\b"])
        fuel_col = find_col(cols, [r"\bfuel\b", r"source"])
        owner_col = find_col(cols, [r"utility", r"owner", r"company", r"developer"])
        sector_col = find_col(cols, [r"\bsector\b", r"category"])
        cap_col = find_col(cols, [r"installed.*capacity", r"\bcapacity\b", r"\b mw\b"])
        if station_col is None or state_col is None: continue
        # Fill down station/state/sector/utility cells where source uses merged cells.
        sub = df.iloc[header_idx+1:].copy()
        for col in [station_col, state_col, owner_col, sector_col, fuel_col]:
            if col is not None and col < sub.shape[1]:
                sub.iloc[:, col] = sub.iloc[:, col].replace("", pd.NA).ffill().fillna("")
        for _, row in sub.iterrows():
            vals = [clean(x) for x in row.tolist()]
            if station_col >= len(vals) or state_col >= len(vals): continue
            name, state = vals[station_col], vals[state_col]
            if not name or not state: continue
            if re.search(r"^(total|grand total|all india|region total|sub total)$", name, re.I): continue
            owner = vals[owner_col] if owner_col is not None and owner_col < len(vals) else ""
            sector = vals[sector_col] if sector_col is not None and sector_col < len(vals) else ""
            raw_fuel = vals[fuel_col] if fuel_col is not None and fuel_col < len(vals) else ""
            capacity = vals[cap_col] if cap_col is not None and cap_col < len(vals) else ""
            # Avoid including industrial/captive establishments even when they say CPP.
            if industrial(name, owner): continue
            sector_l = (sector + " " + owner).lower()
            is_central_state = bool(re.search(r"\b(central|state sector|state government)\b", sector_l))
            is_utility_ipp = bool(re.search(r"\b(utility|ipp|independent power producer)\b", sector_l))
            known_owner = any(x in owner.lower() for x in UTILITY_OWNERS)
            if not (is_central_state or is_utility_ipp or known_owner):
                continue
            records.append({
                "name": name, "state": state, "rawFuel": raw_fuel or source_label,
                "capacity": re.sub(r"[^0-9.]", "", capacity),
                "owner": owner, "year": "", "lat": "", "lon": "",
                "source": f"CEA/NPP {source_label}", "source_as_of": ""
            })
    return records

def main():
    DATA.mkdir(exist_ok=True)
    session = requests.Session()
    session.headers.update(HEADERS)
    response = session.get(REPORTS_URL, timeout=TIMEOUT)
    response.raise_for_status()
    links = fetch_report_links(response.text, response.url)
    if not links:
        raise RuntimeError("Could not locate the expected CEA/NPP Excel report links. Existing inventory was left untouched.")
    collected = []
    report_errors = []
    for label, url in links.items():
        try:
            r = session.get(url, timeout=TIMEOUT)
            r.raise_for_status()

            print(f"DEBUG [{label}] requested URL: {url}")
            print(f"DEBUG [{label}] final URL: {r.url}")
            print(
                f"DEBUG [{label}] content-type: "
                f"{r.headers.get('Content-Type')} | "
                f"bytes: {len(r.content)} | "
                f"signature: {r.content[:8]!r}"
            )

            try:
                workbook = pd.ExcelFile(io.BytesIO(r.content))
                print(f"DEBUG [{label}] sheets: {workbook.sheet_names}")

                for sheet in workbook.sheet_names[:2]:
                    preview = pd.read_excel(
                        workbook,
                        sheet_name=sheet,
                        header=None,
                        nrows=8,
                        dtype=str
                    ).fillna("")

                    for row_no, values in enumerate(
                        preview.itertuples(index=False, name=None)
                    ):
                        print(
                            f"DEBUG [{label}] {sheet} row {row_no + 1}: "
                            f"{[clean(x) for x in values[:12]]}"
                        )

            except Exception as debug_error:
                print(
                    f"DEBUG [{label}] workbook preview failed: "
                    f"{debug_error!r}"
                )

            rows = parse_workbook(r.content, label)
            if rows: collected.extend(rows)
            else: report_errors.append(f"{label}: no station/state table detected")
        except Exception as exc:
            report_errors.append(f"{label}: {exc}")
    if not collected:
        raise RuntimeError("No usable CEA station records were parsed. Existing inventory was left untouched. " + "; ".join(report_errors))

    existing = {}
    if CSV_PATH.exists():
        with CSV_PATH.open("r", encoding="utf-8-sig", newline="") as f:
            for row in csv.DictReader(f):
                existing[(key(row.get("name")), key(row.get("state")))] = row
    added = []
    for row in collected:
        k = (key(row["name"]), key(row["state"]))
        if k not in existing:
            existing[k] = row
            added.append(row)
        else:
            # Update non-empty source fields but never erase prior values.
            old = existing[k]
            for field, value in row.items():
                if value and not old.get(field): old[field] = value

    # Write atomically only after at least one source has been parsed successfully.
    fields = ["name","state","rawFuel","capacity","owner","year","lat","lon","source","source_as_of"]
    tmp = CSV_PATH.with_suffix(".tmp")
    with tmp.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for row in sorted(existing.values(), key=lambda x: (x.get("state",""), x.get("name",""))):
            writer.writerow({k: row.get(k, "") for k in fields})
    tmp.replace(CSV_PATH)

    now = datetime.now(timezone.utc).isoformat()
    previous_log = {"history": []}
    if LOG_PATH.exists():
        try:
            import json
            previous_log = json.loads(LOG_PATH.read_text(encoding="utf-8"))
        except Exception: pass
    entry = {
        "checked_at_utc": now,
        "sources_checked": links,
        "records_seen_in_reports": len(collected),
        "new_records_added": len(added),
        "new_records": added[:250],
        "notes": report_errors
    }
    previous_log.setdefault("history", []).insert(0, entry)
    previous_log["history"] = previous_log["history"][:100]
    LOG_PATH.write_text(json.dumps(previous_log, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"CEA monitor completed. Parsed {len(collected)} source rows; added {len(added)} new station records.")
    for note in report_errors: print("WARNING:", note)

if __name__ == "__main__":
    main()
