"""Application configuration and path memory.

Two files in config/:
  - state.json: session paths, window geometry, active config name
  - <name>.json (e.g. OPPD.json): job-specific keywords
"""

import json
import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional


# Keys stored in state.json. Everything else lives in the job config file.
STATE_KEYS = {
    "active_config",
    "last_existing_folder_path",
    "last_proposed_folder_path",
    "excel_file_path",
}


def _get_project_root() -> str:
    """Get project root (parent of src/)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return str(Path(__file__).resolve().parents[2])


def _get_config_dir() -> str:
    """Get config directory path (project_root/config)."""
    return os.path.join(_get_project_root(), "config")


def _get_state_path() -> str:
    """Path to config/state.json."""
    return os.path.join(_get_config_dir(), "state.json")


def get_active_config_name() -> str:
    """Return the active config name from state.json. Default: OPPD."""
    state = _load_state()
    return state.get("active_config", "OPPD")


def get_available_configs() -> List[str]:
    """List available config names (without .json), excluding state and underscore files."""
    config_dir = _get_config_dir()
    if not os.path.isdir(config_dir):
        return ["OPPD"]
    configs = []
    for f in os.listdir(config_dir):
        if f.endswith(".json") and not f.startswith("_") and f != "state.json":
            configs.append(f[:-5])
    return sorted(configs) if configs else ["OPPD"]


# ---------- State persistence ----------

_STATE_DEFAULTS: Dict[str, Any] = {
    "active_config": "OPPD",
    "last_existing_folder_path": "",
    "last_proposed_folder_path": "",
    "excel_file_path": "",
}


def _load_state() -> Dict[str, Any]:
    """Load state from config/state.json."""
    path = _get_state_path()
    state = dict(_STATE_DEFAULTS)
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            state.update({k: v for k, v in data.items() if k in STATE_KEYS})
    except Exception as e:
        print(f"Error loading state: {e}")

    # Migrate: if _active.json exists, absorb its value and delete it
    active_path = os.path.join(_get_config_dir(), "_active.json")
    try:
        if os.path.exists(active_path):
            with open(active_path, "r", encoding="utf-8") as f:
                active_data = json.load(f)
            if "active_config" in active_data:
                state["active_config"] = active_data["active_config"]
            os.remove(active_path)
            _save_state(state)
    except Exception:
        pass

    return state


def _save_state(state: Dict[str, Any]) -> None:
    """Save state to config/state.json."""
    path = _get_state_path()
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(state, f, indent=2)
    except Exception as e:
        print(f"Error saving state: {e}")


# ---------- Config manager ----------

class PPLXConfigManager:
    """Manages job configuration and session state."""

    def __init__(self, config_name: Optional[str] = None):
        self.state = _load_state()
        self.config_name = config_name or self.state.get("active_config", "OPPD")
        self.config_file = self._find_config_file()
        self.config = self._load_config()

    def _find_config_file(self) -> str:
        """Find config file in config/ folder."""
        config_dir = _get_config_dir()
        config_path = os.path.join(config_dir, f"{self.config_name}.json")

        if getattr(sys, "frozen", False):
            exe_dir = os.path.dirname(sys.executable)
            for p in [
                os.path.join(exe_dir, "config", f"{self.config_name}.json"),
                config_path,
            ]:
                if os.path.exists(p):
                    return p

        os.makedirs(config_dir, exist_ok=True)
        return config_path

    def _load_config(self) -> Dict[str, Any]:
        """Load job configuration from JSON file."""
        default = {}
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, "r", encoding="utf-8") as f:
                    config = json.load(f)
                # Strip legacy keys if present
                for legacy in ("configurations", "selected_config", "ignore_scid_keywords", "0", "1"):
                    config.pop(legacy, None)
                return config
        except Exception as e:
            print(f"Error loading config: {e}")
        return default

    def save_config(self) -> None:
        """Save job configuration to JSON file."""
        try:
            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=2)
        except Exception as e:
            print(f"Error saving config: {e}")

    def get(self, key: str, default: Any = None) -> Any:
        """Get value from state or config."""
        if key in STATE_KEYS:
            return self.state.get(key, default)
        return self.config.get(key, default)

    def set(self, key: str, value: Any) -> None:
        """Set value and persist to the appropriate file."""
        if key in STATE_KEYS:
            self.state[key] = value
            _save_state(self.state)
        else:
            self.config[key] = value
            self.save_config()

    def switch_config(self, name: str) -> None:
        """Switch to a different job config and reload."""
        self.config_name = name
        self.config_file = self._find_config_file()
        self.config = self._load_config()
        self.state["active_config"] = name
        _save_state(self.state)

