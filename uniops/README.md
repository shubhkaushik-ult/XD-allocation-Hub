# UniOps — Centralized Multi-Project Operations & MCP Hub

UniOps is the unified operational backbone for the **XD Allocation**, **Mas SO**, **CUS SO**, **CRON**, and **WMS** ecosystem.

It provides:
1. **Dynamic Configuration via MCP** — Add or update cities, sheet IDs, thresholds, and parameters without hardcoded code changes.
2. **Context Graph with Graft (`trailhq/Graft`)** — Structured codebase dependency tracking for AI coding agents (Claude Code, Cursor, Antigravity).
3. **Automated Error Knowledge Base (KB)** — Root-cause fixes and remediation playbooks.
4. **Centralized Event Telemetry & Audit Logs** — Tracks every pipeline run, config change, and error.

---

## 🛠 Directory Structure

```
uniops/
├── uniops_config.json      # Dynamic configuration store (cities, settings, projects)
├── uniops_kb.json          # Known error patterns and auto-fix playbooks
├── uniops_events.jsonl     # Centralized event stream and audit logs
├── core.py                 # Core Python SDK for all ecosystem tools
├── mcp_server.py           # Model Context Protocol (MCP) server for AI assistants
└── README.md               # Documentation and integration guide
```

---

## 🤖 Using with Claude Code / Cursor / Antigravity

The `.mcp.json` at the root of the workspace automatically registers `uniops` as an active MCP server.

### Adding a New City via Claude Code / MCP:
You can prompt your AI agent:
> *"Add Pune to XD Allocation using tab prefix 'pun', Excel upload mode, and Pune Indent Plan tab."*

The AI will call:
```json
{
  "project": "xd_allocation",
  "city_key": "pune",
  "label": "Pune",
  "tab_prefix": "pun",
  "sheet_id": "excel_upload_only",
  "indent_plan_tab": "Pune Indent Plan",
  "indent_plan_fsn_col": "Pune FSN"
}
```
All running instances of XD Allocation immediately detect the new city dynamically without touching Python code!

### MCP Tools Available:
| Tool Name | Purpose |
| :--- | :--- |
| `uniops_get_config` | Read live config for any project |
| `uniops_list_cities` | List all active cities and settings |
| `uniops_add_city` | Dynamically onboard a new city |
| `uniops_remove_city` | Disable/remove a city |
| `uniops_update_config` | Update arbitrary settings (e.g. batch size) |
| `uniops_log_event` | Log audit trail or execution events |
| `uniops_query_logs` | Query recent logs across projects |
| `uniops_query_kb` | Check known fixes for error messages |
| `uniops_record_fix` | Save new playbook remediation |
| `uniops_health_check` | Verify system health and project count |

---

## 🌲 Graft Context Graph (`trailhq/Graft`)

UniOps includes `graft.json` at the workspace root.

To build and query the codebase graph:
```bash
# Initialize / build graft graph
graft build

# Check callers/callees or blast radius
graft callers xd_processor.py
```

---

## 🐍 Python SDK Integration Example

```python
from uniops.core import uniops_instance

# 1. Fetch dynamic city configuration
cities = uniops_instance.get_cities("xd_allocation")

# 2. Log an execution event
uniops_instance.log(
    level="INFO",
    project="xd_allocation",
    message="Processed PO allocation for Trichy",
    details={"records": 150, "city": "trichy"}
)

# 3. Check KB on error
fix = uniops_instance.find_kb_fix("Tab Tri_2026-10-01_PO not found", project="xd_allocation")
if fix:
    print("Suggested remediation:", fix["remediation_steps"])
```
