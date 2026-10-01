"""
UniOps Core SDK
Centralized configuration, telemetry/logging, and automated knowledge-base registry.
Designed for XD Allocation, Mas SO, CUS SO, CRON, and WMS.
"""

import os
import json
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

# Default paths
DEFAULT_BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = DEFAULT_BASE_DIR / "uniops_config.json"
KB_PATH = DEFAULT_BASE_DIR / "uniops_kb.json"
EVENTS_LOG_PATH = DEFAULT_BASE_DIR / "uniops_events.jsonl"


class UniOps:
    """Core SDK for managing dynamic configuration, audit logs, and knowledge base remediation."""

    def __init__(self, base_dir: Optional[Path] = None):
        self.base_dir = base_dir or DEFAULT_BASE_DIR
        self.config_path = self.base_dir / "uniops_config.json"
        self.kb_path = self.base_dir / "uniops_kb.json"
        self.events_path = self.base_dir / "uniops_events.jsonl"
        self._cached_config: Optional[Dict[str, Any]] = None

    # -------------------------------------------------------------
    # CONFIGURATION MANAGEMENT
    # -------------------------------------------------------------
    def get_config(self, project: Optional[str] = None, force_reload: bool = False) -> Dict[str, Any]:
        """Fetch current configuration. If project is provided, returns that project's config."""
        if not self._cached_config or force_reload:
            if not self.config_path.exists():
                return {}
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    self._cached_config = json.load(f)
            except Exception as e:
                self.log("ERROR", "uniops", f"Failed to load config file: {e}", event_type="CONFIG_READ_ERROR")
                return {}

        if project:
            return self._cached_config.get("projects", {}).get(project, {})
        return self._cached_config

    def get_cities(self, project: str = "xd_allocation") -> Dict[str, Dict[str, Any]]:
        """Helper to get cities mapping for a project (e.g., xd_allocation)."""
        proj = self.get_config(project)
        return proj.get("cities", {})

    def update_config(self, project: str, key_path: str, value: Any, comment: str = "") -> bool:
        """
        Update a configuration property dynamically.
        key_path can be dot-separated (e.g. 'cities.pune.label' or 'settings.batch_size').
        """
        config = self.get_config(force_reload=True)
        if "projects" not in config:
            config["projects"] = {}
        if project not in config["projects"]:
            config["projects"][project] = {"name": project, "settings": {}, "cities": {}}

        # Navigate nested path
        keys = key_path.split(".")
        current = config["projects"][project]
        for k in keys[:-1]:
            if k not in current or not isinstance(current[k], dict):
                current[k] = {}
            current = current[k]
        current[keys[-1]] = value

        config["last_updated"] = datetime.now(timezone.utc).isoformat()

        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(config, f, indent=2)
            self._cached_config = config
            self.log(
                "INFO",
                project,
                f"Configuration updated for '{key_path}'",
                details={"key_path": key_path, "value": value, "comment": comment},
                event_type="CONFIG_UPDATE"
            )
            return True
        except Exception as e:
            self.log("ERROR", project, f"Failed to write config: {e}", event_type="CONFIG_WRITE_ERROR")
            return False

    def add_city(self, project: str, city_key: str, city_data: Dict[str, Any], comment: str = "") -> bool:
        """Add or update a city entry in a project config."""
        city_key = city_key.lower().strip()
        return self.update_config(project, f"cities.{city_key}", city_data, comment=comment or f"Added/Updated city {city_key}")

    def remove_city(self, project: str, city_key: str, comment: str = "") -> bool:
        """Remove or disable a city from a project."""
        city_key = city_key.lower().strip()
        config = self.get_config(force_reload=True)
        try:
            cities = config.get("projects", {}).get(project, {}).get("cities", {})
            if city_key in cities:
                del cities[city_key]
                config["last_updated"] = datetime.now(timezone.utc).isoformat()
                with open(self.config_path, "w", encoding="utf-8") as f:
                    json.dump(config, f, indent=2)
                self._cached_config = config
                self.log("INFO", project, f"Removed city '{city_key}'", details={"city_key": city_key, "comment": comment}, event_type="CONFIG_REMOVE_CITY")
                return True
            return False
        except Exception as e:
            self.log("ERROR", project, f"Failed to remove city '{city_key}': {e}", event_type="CONFIG_WRITE_ERROR")
            return False

    # -------------------------------------------------------------
    # LOGGING & AUDIT TRAIL
    # -------------------------------------------------------------
    def log(self, level: str, project: str, message: str, details: Optional[Dict[str, Any]] = None, event_type: str = "GENERIC") -> None:
        """Append an event log to uniops_events.jsonl."""
        entry = {
            "timestamp": datetime.now(timezone.utc).isoformat(),
            "level": level.upper(),
            "project": project,
            "event": event_type,
            "message": message,
            "details": details or {}
        }
        try:
            self.events_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.events_path, "a", encoding="utf-8") as f:
                f.write(json.dumps(entry) + "\n")
        except Exception:
            pass  # Non-blocking telemetry

    def query_logs(self, project: Optional[str] = None, level: Optional[str] = None, limit: int = 50) -> List[Dict[str, Any]]:
        """Read recent events matching filters."""
        if not self.events_path.exists():
            return []
        results = []
        try:
            with open(self.events_path, "r", encoding="utf-8") as f:
                lines = f.readlines()
            for line in reversed(lines):
                if not line.strip():
                    continue
                try:
                    entry = json.loads(line)
                    if project and entry.get("project") != project:
                        continue
                    if level and entry.get("level") != level.upper():
                        continue
                    results.append(entry)
                    if len(results) >= limit:
                        break
                except Exception:
                    continue
        except Exception as e:
            self.log("ERROR", "uniops", f"Error reading logs: {e}", event_type="LOG_READ_ERROR")
        return results

    # -------------------------------------------------------------
    # KNOWLEDGE BASE & FIX REGISTRY
    # -------------------------------------------------------------
    def find_kb_fix(self, error_text: str, project: Optional[str] = None) -> Optional[Dict[str, Any]]:
        """Look up known fixes in the knowledge base matching the error text."""
        if not self.kb_path.exists():
            return None
        try:
            with open(self.kb_path, "r", encoding="utf-8") as f:
                kb_data = json.load(f)
            items = kb_data.get("knowledge_items", [])
            for item in items:
                item_proj = item.get("project", "global")
                if project and item_proj not in (project, "global"):
                    continue
                pattern = item.get("error_pattern", "")
                if pattern and re.search(pattern, error_text, re.IGNORECASE):
                    return item
        except Exception as e:
            self.log("ERROR", "uniops", f"Error querying KB: {e}", event_type="KB_QUERY_ERROR")
        return None

    def record_kb_fix(self, fix_id: str, error_pattern: str, root_cause: str, remediation_steps: List[str], project: str = "global") -> bool:
        """Register or update a known issue remediation in the KB."""
        try:
            kb_data = {"version": "1.0.0", "knowledge_items": []}
            if self.kb_path.exists():
                with open(self.kb_path, "r", encoding="utf-8") as f:
                    kb_data = json.load(f)
            
            items = kb_data.get("knowledge_items", [])
            # Update if exists
            updated = False
            for idx, item in enumerate(items):
                if item.get("id") == fix_id:
                    items[idx] = {
                        "id": fix_id,
                        "error_pattern": error_pattern,
                        "project": project,
                        "root_cause": root_cause,
                        "remediation_steps": remediation_steps
                    }
                    updated = True
                    break
            if not updated:
                items.append({
                    "id": fix_id,
                    "error_pattern": error_pattern,
                    "project": project,
                    "root_cause": root_cause,
                    "remediation_steps": remediation_steps
                })
            
            kb_data["knowledge_items"] = items
            kb_data["last_updated"] = datetime.now(timezone.utc).isoformat()

            with open(self.kb_path, "w", encoding="utf-8") as f:
                json.dump(kb_data, f, indent=2)

            self.log("INFO", project, f"Registered KB fix {fix_id}", details={"fix_id": fix_id}, event_type="KB_RECORD_FIX")
            return True
        except Exception as e:
            self.log("ERROR", project, f"Failed to record KB fix: {e}", event_type="KB_WRITE_ERROR")
            return False


# Global default instance
uniops_instance = UniOps()
