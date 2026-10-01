# UniOps Operations & XD Allocation Ecosystem Connector

This document is the unified context connector for Claude across the **XD Allocation**, **Mas SO**, **CUS SO**, **CRON**, and **WMS** supply chain ecosystem.

---

## 🏗 Architecture & Dynamic Configuration

- **Core Dynamic Configuration**: `uniops/uniops_config.json`
  - City definitions, sheet IDs, and tool settings are managed dynamically.
  - Never hardcode new cities in `xd_processor.py`. Always add or query via UniOps.
- **Knowledge Base & Error Playbooks**: `uniops/uniops_kb.json`
  - Contains root-cause resolutions for Google Sheet 403 errors, missing FSN mappings, and date tab mismatches.
- **Audit Logs Stream**: `uniops/uniops_events.jsonl`
  - Central log stream for all pipeline executions, config adjustments, and agent actions.
- **Code Graph (TrailHQ Graft)**: `graft.json`
  - Codebase dependency tree across all tools.

---

## 🏙 Active City Profiles & Mappings (XD Allocation)

1. **Trichy (TRZ)**
   - Prefix: `Tri` | Indent Tab: `Trichy Indent Plan` | FSN Col: `Trichy FSN`
   - Sheet ID: `1lXWT-9x3WbNpYYXLZcP4_DEwfuAl1VWbWpouNfTzwXU`
2. **Chennai (MAA)**
   - Prefix: `che` | Indent Tab: `Chennai Indent Plan` | FSN Col: `Chennai FSN`
   - Sheet ID: `1BquGJJri6WpJsUIre7JZLlpOeCR-etR3hLHYbHeOBOo`
3. **Coimbatore (CJB)**
   - Prefix: `coi` | Indent Tab: `Coimbatore Indent Plan` | FSN Col: `Coimbatore FSN`
   - Sheet ID: `19YLdB0JeEnTWEnvmVFVZIlc4D7T0eB1jRXotvSY6jJ8`
4. **Bengaluru (BLR)**
   - Prefix: `ben` | Indent Tab: `Bangalore Indent Plan` | FSN Col: `Bangalore FSN`
   - Sheet ID: `1F2S90yz-rHIDCfAAVXi4p0tQBuDO5fRlm-5T8lzlMFw`
   - Staples Supplier/OUID: Auto-assigned as `OU83946700`
5. **Mumbai (BOM)**
   - Prefix: `mum` | Indent Tab: `Mumbai Indent Plan` | FSN Col: `Mumbai FSN`
   - Sheet ID: `1y2LaBblsqRLX1lG0XpOTVzUGvE40q_z_ZIcLyxPpl54`
6. **Jaipur (JAI)**
   - Prefix: `jai` | Indent Tab: `Jaipur Indent Plan` | FSN Col: `Jaipur FSN`
   - Mode: Excel Upload Only (`excel_upload_only`)

---

## 🛠 Standard Operations & Commands

### Adding / Updating a City in UniOps:
```python
from uniops.core import uniops_instance
uniops_instance.add_city(
    project="xd_allocation",
    city_key="pune",
    city_data={
        "label": "Pune",
        "tab_prefix": "pun",
        "sheet_id": "excel_upload_only",
        "indent_plan_tab": "Pune Indent Plan",
        "indent_plan_fsn_col": "Pune FSN",
        "enabled": True
    },
    comment="Onboarded Pune"
)
```

### Checking System Health & Log Stream:
```bash
python uniops/mcp_server.py --status
python uniops/mcp_server.py --list-cities
```

### Local Development & Testing:
- Start web interface: `python xd_processor.py`
- Run local tests: `python -m unittest discover`
- Deploy to Vercel: `vercel --prod`
