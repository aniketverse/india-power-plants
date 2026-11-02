# Aniketverse — India Electricity Power Stations

This repository contains the public dashboard and a scheduled CEA/National Power Portal (NPP) source monitor.

## What updates automatically

- GitHub Actions runs once daily (07:55 IST) and can also be run manually.
- It checks the official NPP published reports page for daily station-wise generation, monthly station-wise details and installed-capacity reports.
- It parses eligible station/state records, screens industrial names and owners, and reconciles new records into `data/cea_additions.csv`.
- It records each run and new additions in `data/update_log.json`.
- Existing records are never deleted merely because a station is absent from one daily report.
- New CEA records without verified coordinates are shown in the inventory but are not plotted on the map until coordinates are available.

## Important data-quality rules

1. Historical WRI GPPD v1.3.0 remains the map's baseline and is not a current official register.
2. CEA/NPP is used as the scheduled official source for new station-name detection.
3. Central and State sector station records are eligible for automatic addition.
4. Private-sector records are included only when the report identifies a utility/IPP or the owner matches the utility allow-list in `scripts/update_data.py`.
5. Industrial/captive establishments are excluded using station-name and owner-name rules. This is a conservative automated screen, not a legal certification; review the `update_log.json` and refine `INDUSTRY` / `UTILITY_OWNERS` when official names change.
6. ISRO–NITI Energy Map is provided as a geographic cross-check. Its own disclaimer says third-party data may be incomplete.

## Enable the website

1. In the GitHub repository, open **Settings → Pages**.
2. Under **Build and deployment**, choose **Deploy from a branch**.
3. Select branch **main** and folder **/(root)**, then click **Save**.
4. Wait a minute or two. Your site address will be:
   `https://aniketverse.github.io/india-power-plants/`

## Enable automatic updates

1. Open the repository's **Actions** tab.
2. If GitHub asks to enable workflows, click **I understand my workflows, go ahead and enable them**.
3. Open **Refresh India power station inventory**.
4. Click **Run workflow** once to test it. The daily schedule then runs automatically.

## Sources

- National Power Portal published reports: https://npp.gov.in/publishedReports
- CEA daily generation reports: https://cea.nic.in/reports/external/daily-generation?lang=en
- ISRO–NITI Aayog Energy Map: https://vedas.sac.gov.in/energymap/view/powergis.jsp
- WRI GPPD historical baseline: https://github.com/wri/global-power-plant-database
- ListBase India GPPD mirror: https://www.listbase.org/en/environment/in-power-plants/

## Limitations

The CEA report pages are published as Excel/PDF files, not a stable public API. If CEA changes its page layout or workbook columns, the monitor fails safely and leaves the previous inventory unchanged. The current monitor is strongest for conventional grid-connected station names; renewable project-level coverage and precise geocoding require separate structured source feeds. This project must not be described as a guaranteed exhaustive, real-time national register.
