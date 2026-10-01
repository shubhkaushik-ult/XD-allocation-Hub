# Project Changes Log & Version History

## Version 2.6.0 (2026-09-30)

### Features & Adjustments

1. **UniOps Ecosystem & Dynamic MCP Integration**:
   - Integrated UniOps central configuration (`uniops/uniops_config.json`) and core SDK (`uniops/core.py`).
   - Removed hardcoded city dictionary dependencies from `xd_processor.py` to allow dynamic scaling and city onboarding via MCP.
   - Added `uniops/mcp_server.py` exposing MCP tools for dynamic config management, centralized audit logging (`uniops/uniops_events.jsonl`), and automated KB fix resolution (`uniops/uniops_kb.json`).
   - Added **TrailHQ Graft** context graph configuration (`graft.json`) and `.mcp.json` for AI coding agents (Claude Code, Cursor, Antigravity).

---

## Version 2.5.0 (2026-09-28)

### Features & Adjustments

1. **Jaipur (JAI) City Integration**:
   - Added support for Jaipur in the XD Allocation Hub.
   - Configured the workflow to process raw Excel allocation files exclusively, bypassing Google Sheets fetches.
   - Updated the web UI grid to feature the new Jaipur active card.

---

## Version 2.4.0 (2026-09-17)

### Features & Adjustments

1. **Bengaluru (BLR) Staples OU ID Automation**:
   - Updated `process_addon_df` in `xd_processor.py` to accept an `is_staples: bool = False` flag.
   - For items uploaded via the Staples upload option in Bengaluru (`city_key in ["bengaluru", "blr"]`), `Supplier ID` and `OUID` are automatically assigned as **`OU83946700`**.
   - Decoupled `addon_df` and `staples_df` parsing in `build_excel_report` so that Add-on uploads (`is_staples=False`) and Staples uploads (`is_staples=True`) are evaluated independently.

2. **Project Documentation Suite**:
   - Established the `/docs` documentation repository containing system architecture details, change logs, log/troubleshooting procedures, and index README.

---

## Version 2.3.0 (Prior Release)
- Integrated Grocer SO/PO generator (`gro_so_script.py`).
- Added support for multi-city ZIP downloads (`ALL` cities selection).
- Implemented client-side file compression and multi-file drag & drop support in UI.
- Unified column coalescing across raw indents, add-ons, and staples files.
