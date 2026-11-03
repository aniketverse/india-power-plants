#!/usr/bin/env python3
"""Refresh the dashboard's CEA station additions feed from the official NPP DGR-2 report.

The updater intentionally includes every Indian power station listed in the CEA
station-wise report, including private-sector, industrial and captive stations.
It does not infer stations that are absent from the CEA report.

Safety behaviour:
- Selects only Daily Report 02 (station-wise), never the All-India summary.
- Validates the downloaded Excel workbook and expected CEA report structure.
- Requires a reasonable number of parsed station rows before writing anything.
- Keeps existing stations if a later report omits them.
- Updates CEA-sourced fields for stations already present.
- Writes the CSV atomically; a failed run leaves the existing file untouched.
"""

from __future__ import annotations

import csv
import io
import json
import math
import re
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin, urlsplit

import pandas as pd
import requests
from bs4 import BeautifulSoup

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
CSV_PATH = DATA_DIR / "cea_additions.csv"
LOG_PATH = DATA_DIR / "update_log.json"
REPORTS_URL = "https://npp.gov.in/publishedReports"
TIMEOUT_SECONDS = 60
MIN_EXPECTED_STATIONS = 100
USER_AGENT = "Aniketverse-India-Power-Plants/2.0 (CEA public-report inventory updater)"

# CEA uses uppercase state labels in its JasperReports workbook.
STATE_MAP = {
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
    "JAMMU & KASHMIR": "Jammu and Kashmir",
    "LADAKH": "Ladakh",
    "DELHI": "Delhi",
    "ANDAMAN AND NICOBAR ISLANDS": "Andaman and Nicobar Islands",
    "CHANDIGARH": "Chandigarh",
    "PUDUCHERRY": "Puducherry",
    "PONDICHERRY": "Puducherry",
    "DADRA AND NAGAR HAVELI AND DAMAN AND DIU": "Dadra and Nagar Haveli and Daman and Diu",
}

CSV_FIELDS = [
    "name", "state", "rawFuel", "capacity", "owner", "year", "lat", "lon",
    "source", "source_as_of",
]


def clean(value: object) -> str:
    """Convert an Excel cell to a trimmed, whitespace-normalised string."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = str(value).replace("\xa0", " ")
    return re.sub(r"\s+", " ", text).strip()


def normal_key(value: object) -> str:
    return re.sub(r"[^a-z0-9]", "", clean(value).lower())


def parse_number(value: object) -> float | None:
    """Read CEA numeric cells; zero is valid and must not be discarded."""
    text = clean(value).replace(",", "").replace(" ", "")
    if not text or text in {"-", "—", "NA", "N/A"}:
        return None
    try:
        number = float(text)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(number) or number < 0:
        return None
    return number


def format_number(number: float) -> str:
    return format(number, ".12g")


def fetch_latest_dgr2_link(html: str, base_url: str) -> str:
    """Return the newest XLS/XLSX URL for CEA Daily Generation Report 02 only."""
    soup = BeautifulSoup(html, "html.parser")
    candidates: list[tuple[str, str]] = []

    for anchor in soup.find_all("a", href=True):
        url = urljoin(base_url, anchor["href"])
        path = urlsplit(url).path.lower()

        # dgr1 is the all-India overview; dgr2 is the station-wise report.
        match = re.search(r"/dgr2-(\d{4}-\d{2}-\d{2})\.xlsx?$", path)
        if match:
            candidates.append((match.group(1), url))

    if not candidates:
        raise RuntimeError(
            "Could not find the CEA Daily Generation Report 02 (DGR-2) Excel link. "
            "No inventory files were changed."
        )

    candidates.sort(key=lambda item: item[0], reverse=True)
    return candidates[0][1]


def read_excel_workbook(content: bytes) -> dict[str, pd.DataFrame]:
    """Open legacy XLS or XLSX using the matching installed pandas engine."""
    if content.startswith(bytes.fromhex("D0CF11E0A1B11AE1")):
        engine = "xlrd"
    elif content.startswith(b"PK\x03\x04"):
        engine = "openpyxl"
    else:
        raise ValueError(
            "The downloaded report is not recognised as an XLS/XLSX file "
            f"(first 8 bytes: {content[:8]!r})."
        )

    excel = pd.ExcelFile(io.BytesIO(content), engine=engine)
    return {
        sheet_name: pd.read_excel(
            excel, sheet_name=sheet_name, header=None, dtype=object
        ).fillna("")
        for sheet_name in excel.sheet_names
    }


def find_report_header(df: pd.DataFrame) -> tuple[int, int, int] | None:
    """Locate the header row and the station/capacity columns by their labels."""
    for row_index in range(min(len(df), 35)):
        cells = [clean(value).upper() for value in df.iloc[row_index].tolist()]
        station_col = next(
            (i for i, value in enumerate(cells) if "POWER STATION" in value), None
        )
        capacity_col = next(
            (
                i for i, value in enumerate(cells)
                if "MONITORED" in value and "CAP" in value
            ),
            None,
        )
        if station_col is not None and capacity_col is not None:
            return row_index, station_col, capacity_col
    return None


def extract_report_date(workbook: dict[str, pd.DataFrame]) -> str:
    """Extract 'REPORT FOR dd/mm/yyyy' from the CEA report heading."""
    top_text_parts: list[str] = []
    for df in workbook.values():
        for value in df.head(12).to_numpy().ravel():
            text = clean(value)
            if text:
                top_text_parts.append(text)
    top_text = " ".join(top_text_parts)

    match = re.search(
        r"REPORT\s+FOR\s+(\d{1,2}[/-]\d{1,2}[/-]\d{4})",
        top_text,
        re.IGNORECASE,
    )
    if not match:
        raise ValueError("The workbook did not contain a recognisable CEA report date.")
    date_text = match.group(1)
    date_format = "%d/%m/%Y" if "/" in date_text else "%d-%m-%Y"
    return datetime.strptime(date_text, date_format).date().isoformat()


def parse_dgr2_workbook(content: bytes) -> tuple[list[dict[str, str]], str, str]:
    """Parse CEA DGR-2 state/sector/type hierarchy into station-level rows."""
    workbook = read_excel_workbook(content)
    report_date = extract_report_date(workbook)
    collected: list[dict[str, str]] = []
    selected_sheet = ""

    for sheet_name, df in workbook.items():
        header = find_report_header(df)
        if header is None:
            continue

        header_row, station_col, capacity_col = header
        selected_sheet = sheet_name
        current_state = ""
        current_sector = ""
        current_fuel = ""

        for row_values in df.iloc[header_row + 1 :].itertuples(index=False, name=None):
            values = [clean(value) for value in row_values]
            if station_col >= len(values):
                continue

            station_label = values[station_col]
            upper = station_label.upper().strip()
            if not upper:
                continue

            # CEA grouping rows: Region -> State -> Sector -> Type -> Stations.
            if upper in STATE_MAP:
                current_state = STATE_MAP[upper]
                current_sector = ""
                current_fuel = ""
                continue

            if upper.startswith("SECTOR:"):
                current_sector = next(
                    (value for value in values[station_col + 1 :] if value), ""
                )
                current_fuel = ""
                continue

            if upper.startswith("TYPE:"):
                current_fuel = next(
                    (value for value in values[station_col + 1 :] if value), ""
                )
                continue

            # Discard region/state summaries, headings, unit-level detail and footnotes.
            if not current_state:
                continue
            if upper.startswith("UNIT"):
                continue
            if upper in {"STATE TOTAL", "REGION TOTAL", "TOTAL", "GRAND TOTAL"}:
                continue
            if "TOTAL" in upper or upper.startswith("REMARK") or upper.startswith("REPORT VERSION"):
                continue
            if upper.startswith("SUBREPORT"):
                continue

            if capacity_col >= len(values):
                continue
            capacity = parse_number(values[capacity_col])
            if capacity is None:
                continue

            # No sector, owner-name, industrial-name or captive-plant exclusions.
            collected.append(
                {
                    "name": station_label,
                    "state": current_state,
                    "rawFuel": current_fuel or "Other",
                    "capacity": format_number(capacity),
                    "owner": "",
                    "year": "",
                    "lat": "",
                    "lon": "",
                    "source": "CEA/NPP Daily Station-wise Generation Report (DGR-2)",
                    "source_as_of": report_date,
                }
            )

    if not selected_sheet:
        raise ValueError(
            "No worksheet contained both POWER STATION and MONITORED CAP. headers. "
            "The CEA workbook layout may have changed."
        )

    # Collapse repeated station labels in the same state without dropping state differences.
    unique: dict[tuple[str, str], dict[str, str]] = {}
    for row in collected:
        unique[(normal_key(row["name"]), normal_key(row["state"]))] = row

    return list(unique.values()), report_date, selected_sheet


def read_existing_inventory() -> dict[tuple[str, str], dict[str, str]]:
    existing: dict[tuple[str, str], dict[str, str]] = {}
    if not CSV_PATH.exists() or CSV_PATH.stat().st_size == 0:
        return existing

    with CSV_PATH.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for source_row in reader:
            row = {field: clean(source_row.get(field, "")) for field in CSV_FIELDS}
            if not row["name"] or not row["state"]:
                continue
            existing[(normal_key(row["name"]), normal_key(row["state"]))] = row
    return existing


def write_inventory_atomically(rows: dict[tuple[str, str], dict[str, str]]) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    temp_path = CSV_PATH.with_suffix(".tmp")
    with temp_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=CSV_FIELDS)
        writer.writeheader()
        for row in sorted(
            rows.values(), key=lambda item: (item.get("state", ""), item.get("name", ""))
        ):
            writer.writerow({field: row.get(field, "") for field in CSV_FIELDS})
    temp_path.replace(CSV_PATH)


def update_log(
    report_url: str,
    report_date: str,
    parsed_count: int,
    added_rows: list[dict[str, str]],
    refreshed_count: int,
) -> None:
    previous = {"history": []}
    if LOG_PATH.exists():
        try:
            parsed = json.loads(LOG_PATH.read_text(encoding="utf-8"))
            if isinstance(parsed, dict):
                previous = parsed
        except (OSError, json.JSONDecodeError):
            pass

    history = previous.get("history", [])
    if not isinstance(history, list):
        history = []

    history.insert(
        0,
        {
            "checked_at_utc": datetime.now(timezone.utc).isoformat(),
            "sources_checked": {"daily_stationwise": report_url},
            "report_as_of": report_date,
            "records_seen_in_reports": parsed_count,
            "new_records_added": len(added_rows),
            "existing_records_refreshed": refreshed_count,
            "new_records": added_rows[:250],
            "notes": [
                "Source is CEA/NPP Daily Generation Report 02 (station-wise).",
                "All stations listed by the source are included regardless of sector or industry.",
            ],
        },
    )
    previous["history"] = history[:100]

    temp_path = LOG_PATH.with_suffix(".tmp")
    temp_path.write_text(
        json.dumps(previous, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    temp_path.replace(LOG_PATH)


def main() -> None:
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT})

    print(f"Reading official CEA/NPP report index: {REPORTS_URL}")
    page = session.get(REPORTS_URL, timeout=TIMEOUT_SECONDS)
    page.raise_for_status()

    report_url = fetch_latest_dgr2_link(page.text, page.url)
    print(f"Selected station-wise DGR-2 report: {report_url}")

    response = session.get(report_url, timeout=TIMEOUT_SECONDS)
    response.raise_for_status()
    print(
        "Downloaded report: "
        f"{len(response.content):,} bytes; "
        f"Content-Type={response.headers.get('Content-Type', 'unknown')}"
    )

    parsed_rows, report_date, sheet_name = parse_dgr2_workbook(response.content)
    print(
        f"Parsed {len(parsed_rows)} unique station rows from sheet '{sheet_name}' "
        f"(report date: {report_date})."
    )

    if len(parsed_rows) < MIN_EXPECTED_STATIONS:
        raise RuntimeError(
            f"Only {len(parsed_rows)} station rows were parsed; expected at least "
            f"{MIN_EXPECTED_STATIONS}. Existing inventory was left untouched."
        )

    existing = read_existing_inventory()
    added: list[dict[str, str]] = []
    refreshed = 0

    # These CEA fields are refreshed on every run. Coordinates and other manually
    # enriched fields are retained from the existing inventory.
    source_fields = ["rawFuel", "capacity", "source", "source_as_of"]

    for row in parsed_rows:
        identity = (normal_key(row["name"]), normal_key(row["state"]))
        if identity not in existing:
            existing[identity] = row.copy()
            added.append(row.copy())
        else:
            old = existing[identity]
            changed = False
            for field in source_fields:
                new_value = row.get(field, "")
                if new_value and old.get(field, "") != new_value:
                    old[field] = new_value
                    changed = True
            if changed:
                refreshed += 1

    # Write only after successful download, structure validation and row-count check.
    write_inventory_atomically(existing)
    update_log(report_url, report_date, len(parsed_rows), added, refreshed)

    print(
        "CEA inventory refresh completed successfully. "
        f"Parsed: {len(parsed_rows)} | New stations: {len(added)} | "
        f"Existing station records refreshed: {refreshed} | "
        f"Total feed rows: {len(existing)}"
    )


if __name__ == "__main__":
    main()
