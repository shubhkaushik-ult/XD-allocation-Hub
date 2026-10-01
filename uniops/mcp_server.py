"""
UniOps MCP Server & HTTP Connector
Exposes tools to AI coding assistants (Claude Web, Claude Desktop, Claude Code, Cursor, Antigravity)
for dynamic multi-project configuration, city onboarding, centralized logging, and KB queries.
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Any, Dict, List, Optional

# Add parent directory to path so uniops can be imported
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from uniops.core import UniOps, uniops_instance

# Check if FastMCP is available
try:
    from mcp.server.fastmcp import FastMCP
    HAS_FASTMCP = True
except ImportError:
    HAS_FASTMCP = False


# =====================================================================
# TOOL IMPLEMENTATIONS
# =====================================================================

def tool_get_config(project: str = "xd_allocation", key: Optional[str] = None) -> Dict[str, Any]:
    """Retrieve current dynamic configuration for a project or specific key."""
    cfg = uniops_instance.get_config(project, force_reload=True)
    if key and isinstance(cfg, dict):
        return {key: cfg.get(key)}
    return cfg


def tool_list_cities(project: str = "xd_allocation") -> Dict[str, Any]:
    """List all configured cities and their Google Sheet / Excel settings."""
    return uniops_instance.get_cities(project)


def tool_add_city(
    project: str,
    city_key: str,
    label: str,
    tab_prefix: str,
    sheet_id: str = "excel_upload_only",
    indent_plan_tab: Optional[str] = None,
    indent_plan_fsn_col: Optional[str] = None,
    indent_plan_sheet_id: Optional[str] = None,
    comment: str = ""
) -> Dict[str, Any]:
    """
    Dynamically add or scale a new city to any tool without changing code.
    """
    city_data = {
        "label": label,
        "tab_prefix": tab_prefix,
        "sheet_id": sheet_id,
        "indent_plan_tab": indent_plan_tab or f"{label} Indent Plan",
        "indent_plan_fsn_col": indent_plan_fsn_col or f"{label} FSN",
        "enabled": True
    }
    if indent_plan_sheet_id:
        city_data["indent_plan_sheet_id"] = indent_plan_sheet_id

    success = uniops_instance.add_city(project, city_key, city_data, comment=comment or f"Added city {label} via MCP")
    return {
        "success": success,
        "project": project,
        "city_key": city_key,
        "data": city_data,
        "message": f"City '{label}' ({city_key}) successfully {'added' if success else 'failed to add'}."
    }


def tool_remove_city(project: str, city_key: str, comment: str = "") -> Dict[str, Any]:
    """Remove or disable a city from a tool's dynamic configuration."""
    success = uniops_instance.remove_city(project, city_key, comment=comment or f"Removed city {city_key} via MCP")
    return {
        "success": success,
        "project": project,
        "city_key": city_key,
        "message": f"City '{city_key}' {'removed' if success else 'not found or failed to remove'}."
    }


def tool_update_config(project: str, key_path: str, value: Any, comment: str = "") -> Dict[str, Any]:
    """Update any project configuration setting."""
    success = uniops_instance.update_config(project, key_path, value, comment=comment)
    return {
        "success": success,
        "project": project,
        "key_path": key_path,
        "new_value": value
    }


def tool_log_event(level: str, project: str, message: str, details: Optional[Dict[str, Any]] = None, event_type: str = "AGENT_ACTION") -> Dict[str, Any]:
    """Log an operational event, deployment change, or incident to the centralized log stream."""
    uniops_instance.log(level, project, message, details=details, event_type=event_type)
    return {"status": "logged", "level": level, "project": project, "message": message}


def tool_query_logs(project: Optional[str] = None, level: Optional[str] = None, limit: int = 20) -> List[Dict[str, Any]]:
    """Query recent execution and error logs across tools."""
    return uniops_instance.query_logs(project=project, level=level, limit=limit)


def tool_query_kb(error_text: str, project: Optional[str] = None) -> Dict[str, Any]:
    """Search knowledge base for known fixes and root-cause solutions."""
    fix = uniops_instance.find_kb_fix(error_text, project=project)
    if fix:
        return {"matched": True, "knowledge_item": fix}
    return {"matched": False, "message": "No specific KB remediation found for this error pattern."}


def tool_record_fix(fix_id: str, error_pattern: str, root_cause: str, remediation_steps: List[str], project: str = "global") -> Dict[str, Any]:
    """Register a new auto-fix / playbook resolution into the knowledge base."""
    success = uniops_instance.record_kb_fix(fix_id, error_pattern, root_cause, remediation_steps, project=project)
    return {"success": success, "fix_id": fix_id}


def tool_get_latest_updates_and_errors() -> Dict[str, Any]:
    """
    Get live overview of latest updates, recent changes, active errors, and possible issues.
    Specifically designed for real-time awareness in Claude.
    """
    cfg = uniops_instance.get_config(force_reload=True)
    logs = uniops_instance.query_logs(limit=25)
    recent_errors = [l for l in logs if l.get("level") in ("ERROR", "WARNING")]
    
    # Check KB items
    kb_data = {}
    if uniops_instance.kb_path.exists():
        try:
            with open(uniops_instance.kb_path, "r", encoding="utf-8") as f:
                kb_data = json.load(f)
        except Exception:
            pass

    return {
        "status": "LIVE",
        "last_config_update": cfg.get("last_updated"),
        "registered_projects": list(cfg.get("projects", {}).keys()),
        "xd_cities": list(cfg.get("projects", {}).get("xd_allocation", {}).get("cities", {}).keys()),
        "recent_errors_and_warnings": recent_errors,
        "recent_events_count": len(logs),
        "known_issue_playbooks": kb_data.get("knowledge_items", [])
    }


def tool_health_check() -> Dict[str, Any]:
    """Check health and status of UniOps configuration, logs, and registered projects."""
    cfg = uniops_instance.get_config(force_reload=True)
    logs = uniops_instance.query_logs(limit=5)
    return {
        "status": "HEALTHY",
        "registered_projects": list(cfg.get("projects", {}).keys()),
        "xd_cities_count": len(cfg.get("projects", {}).get("xd_allocation", {}).get("cities", {})),
        "recent_events_count": len(logs),
        "version": cfg.get("version", "1.0.0")
    }


# =====================================================================
# MCP TOOL DEFINITIONS & DISPATCHER
# =====================================================================

MCP_TOOLS_SPEC = [
    {
        "name": "uniops_get_latest_updates_and_errors",
        "description": "Fetch real-time updates, latest changes, recent errors, and potential issues across all tools (XD Allocation, Mas SO, CUS SO, CRON, WMS).",
        "inputSchema": {
            "type": "object",
            "properties": {}
        }
    },
    {
        "name": "uniops_get_config",
        "description": "Get live dynamic configuration for any project or specific key",
        "inputSchema": {
            "type": "object",
            "properties": {
                "project": {"type": "string", "default": "xd_allocation"},
                "key": {"type": "string"}
            }
        }
    },
    {
        "name": "uniops_list_cities",
        "description": "List all configured cities for XD Allocation with sheet IDs and tab prefixes",
        "inputSchema": {
            "type": "object",
            "properties": {
                "project": {"type": "string", "default": "xd_allocation"}
            }
        }
    },
    {
        "name": "uniops_add_city",
        "description": "Dynamically add or scale a new city to XD Allocation without editing code",
        "inputSchema": {
            "type": "object",
            "required": ["project", "city_key", "label", "tab_prefix"],
            "properties": {
                "project": {"type": "string"},
                "city_key": {"type": "string"},
                "label": {"type": "string"},
                "tab_prefix": {"type": "string"},
                "sheet_id": {"type": "string", "default": "excel_upload_only"},
                "indent_plan_tab": {"type": "string"},
                "indent_plan_fsn_col": {"type": "string"},
                "indent_plan_sheet_id": {"type": "string"},
                "comment": {"type": "string"}
            }
        }
    },
    {
        "name": "uniops_remove_city",
        "description": "Remove or disable a city from a tool's dynamic configuration",
        "inputSchema": {
            "type": "object",
            "required": ["project", "city_key"],
            "properties": {
                "project": {"type": "string"},
                "city_key": {"type": "string"},
                "comment": {"type": "string"}
            }
        }
    },
    {
        "name": "uniops_update_config",
        "description": "Update any project configuration parameter",
        "inputSchema": {
            "type": "object",
            "required": ["project", "key_path", "value"],
            "properties": {
                "project": {"type": "string"},
                "key_path": {"type": "string"},
                "value": {},
                "comment": {"type": "string"}
            }
        }
    },
    {
        "name": "uniops_query_logs",
        "description": "Query recent execution and error logs across projects",
        "inputSchema": {
            "type": "object",
            "properties": {
                "project": {"type": "string"},
                "level": {"type": "string"},
                "limit": {"type": "integer", "default": 20}
            }
        }
    },
    {
        "name": "uniops_query_kb",
        "description": "Search knowledge base for known fixes, root causes, and remediation playbooks",
        "inputSchema": {
            "type": "object",
            "required": ["error_text"],
            "properties": {
                "error_text": {"type": "string"},
                "project": {"type": "string"}
            }
        }
    },
    {
        "name": "uniops_record_fix",
        "description": "Register a new known fix in the knowledge base",
        "inputSchema": {
            "type": "object",
            "required": ["fix_id", "error_pattern", "root_cause", "remediation_steps"],
            "properties": {
                "fix_id": {"type": "string"},
                "error_pattern": {"type": "string"},
                "root_cause": {"type": "string"},
                "remediation_steps": {"type": "array", "items": {"type": "string"}},
                "project": {"type": "string", "default": "global"}
            }
        }
    },
    {
        "name": "uniops_health_check",
        "description": "Check overall UniOps health and registered tool count",
        "inputSchema": {"type": "object", "properties": {}}
    }
]


def handle_mcp_jsonrpc(req: Dict[str, Any]) -> Dict[str, Any]:
    """Process a single JSON-RPC 2.0 MCP request dict and return the response dict."""
    req_id = req.get("id")
    method = req.get("method")
    params = req.get("params", {})

    if method == "initialize":
        return {
            "jsonrpc": "2.0",
            "id": req_id,
            "result": {
                "protocolVersion": "2024-11-05",
                "capabilities": {
                    "tools": {"listChanged": False},
                    "logging": {}
                },
                "serverInfo": {
                    "name": "uniops-mcp",
                    "version": "1.0.0"
                }
            }
        }
    elif method in ("notifications/initialized", "initialized"):
        return {"jsonrpc": "2.0", "id": req_id, "result": {}}
    elif method == "ping":
        return {"jsonrpc": "2.0", "id": req_id, "result": {}}
    elif method in ("tools/list", "tool/list"):
        return {"jsonrpc": "2.0", "id": req_id, "result": {"tools": MCP_TOOLS_SPEC}}
    elif method in ("tools/call", "tool/call"):
        name = params.get("name")
        args = params.get("arguments", {})
        
        output = None
        if name in ("uniops_get_latest_updates_and_errors", "get_latest_updates_and_errors"):
            output = tool_get_latest_updates_and_errors()
        elif name in ("uniops_get_config", "get_config"):
            output = tool_get_config(args.get("project", "xd_allocation"), args.get("key"))
        elif name in ("uniops_list_cities", "list_cities"):
            output = tool_list_cities(args.get("project", "xd_allocation"))
        elif name in ("uniops_add_city", "add_city"):
            output = tool_add_city(
                project=args.get("project", "xd_allocation"),
                city_key=args.get("city_key", ""),
                label=args.get("label", ""),
                tab_prefix=args.get("tab_prefix", ""),
                sheet_id=args.get("sheet_id", "excel_upload_only"),
                indent_plan_tab=args.get("indent_plan_tab"),
                indent_plan_fsn_col=args.get("indent_plan_fsn_col"),
                indent_plan_sheet_id=args.get("indent_plan_sheet_id"),
                comment=args.get("comment", "")
            )
        elif name in ("uniops_remove_city", "remove_city"):
            output = tool_remove_city(args.get("project", "xd_allocation"), args.get("city_key", ""), args.get("comment", ""))
        elif name in ("uniops_update_config", "update_config"):
            output = tool_update_config(args.get("project", "xd_allocation"), args.get("key_path", ""), args.get("value"), args.get("comment", ""))
        elif name in ("uniops_log_event", "log_event"):
            output = tool_log_event(args.get("level", "INFO"), args.get("project", "general"), args.get("message", ""), args.get("details"), args.get("event_type", "AGENT_ACTION"))
        elif name in ("uniops_query_logs", "query_logs"):
            output = tool_query_logs(args.get("project"), args.get("level"), args.get("limit", 20))
        elif name in ("uniops_query_kb", "query_kb"):
            output = tool_query_kb(args.get("error_text", ""), args.get("project"))
        elif name in ("uniops_record_fix", "record_fix"):
            output = tool_record_fix(args.get("fix_id", ""), args.get("error_pattern", ""), args.get("root_cause", ""), args.get("remediation_steps", []), args.get("project", "global"))
        elif name in ("uniops_health_check", "health_check"):
            output = tool_health_check()
        else:
            output = {"error": f"Unknown tool: {name}"}

        return {
            "jsonrpc": "2.0",
            "id": req_id,
            "result": {
                "content": [{"type": "text", "text": json.dumps(output, indent=2)}]
            }
        }
    else:
        return {"jsonrpc": "2.0", "id": req_id, "result": {}}


def run_stdio_jsonrpc():
    """Stdio JSON-RPC 2.0 loop."""
    for line in sys.stdin:
        if not line.strip():
            continue
        try:
            req = json.loads(line)
            res = handle_mcp_jsonrpc(req)
            sys.stdout.write(json.dumps(res) + "\n")
            sys.stdout.flush()
        except Exception as e:
            err_res = {"jsonrpc": "2.0", "id": None, "error": {"code": -32603, "message": str(e)}}
            sys.stdout.write(json.dumps(err_res) + "\n")
            sys.stdout.flush()


def main():
    parser = argparse.ArgumentParser(description="UniOps MCP Server & CLI")
    parser.add_argument("--status", action="store_true", help="Print UniOps health status")
    parser.add_argument("--list-cities", action="store_true", help="List configured cities")
    parser.add_argument("--jsonrpc", action="store_true", help="Force stdio JSON-RPC loop")
    args = parser.parse_args()

    if args.status:
        print(json.dumps(tool_health_check(), indent=2))
        return

    if args.list_cities:
        print(json.dumps(tool_list_cities(), indent=2))
        return

    run_stdio_jsonrpc()


if __name__ == "__main__":
    main()
