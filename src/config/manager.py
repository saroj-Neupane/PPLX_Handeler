"""Application configuration and path memory."""

import json
import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

# Keys stored in config/state.json (session/preferences). All others go to config/<name>.json.
STATE_KEYS = {
    "last_input_directory",
    "last_output_directory",
    "last_existing_folder_path",
    "last_proposed_folder_path",
    "window_geometry",
    "recent_files",
    "excel_file_path",
    "last_aux_values",
    "auto_fill_aux1",
    "auto_fill_aux2",
    "default_aux_values",
}


def _get_project_root() -> str:
    """Get project root (parent of src/)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return str(Path(__file__).resolve().parents[2])


def _get_config_dir() -> str:
    """Get config directory path (project_root/config)."""
    return os.path.join(_get_project_root(), "config")


def _get_active_config_path() -> str:
    """Path to config/_active.json (replaces root config.json)."""
    return os.path.join(_get_config_dir(), "_active.json")


def _get_state_path() -> str:
    """Path to config/state.json."""
    return os.path.join(_get_config_dir(), "state.json")


def get_available_configs() -> List[str]:
    """List available config names (without .json) in config folder, excluding _active.json."""
    config_dir = _get_config_dir()
    if not os.path.isdir(config_dir):
        return ["OPPD"]
    configs = []
    for f in os.listdir(config_dir):
        if f.endswith(".json") and not f.startswith("_"):
            configs.append(f[:-5])
    return sorted(configs) if configs else ["OPPD"]


def get_active_config_name() -> str:
    """Read active config from config/_active.json. Default: OPPD."""
    path = _get_active_config_path()
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data.get("active_config", "OPPD")
    except Exception:
        pass
    return "OPPD"


def set_active_config(name: str) -> None:
    """Write active config to config/_active.json."""
    path = _get_active_config_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    data = {"active_config": name}
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        print(f"Error saving active config: {e}")


def _load_state() -> Dict[str, Any]:
    """Load state from config/state.json."""
    default = {
        "last_input_directory": "",
        "last_output_directory": "",
        "last_existing_folder_path": "",
        "last_proposed_folder_path": "",
        "window_geometry": "1000x700",
        "recent_files": [],
        "excel_file_path": "",
        "last_aux_values": {},
        "auto_fill_aux1": False,
        "auto_fill_aux2": False,
        "default_aux_values": [""] * 8,
    }
    path = _get_state_path()
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            for k, v in default.items():
                if k not in data:
                    data[k] = v
            return data
    except Exception as e:
        print(f"Error loading state: {e}")
    return default


def _save_state(state: Dict[str, Any]) -> None:
    """Save state to config/state.json."""
    path = _get_state_path()
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(state, f, indent=2)
    except Exception as e:
        print(f"Error saving state: {e}")


class PPLXConfigManager:
    """Manages application configuration and path memory."""

    def __init__(self, config_name: Optional[str] = None):
        """
        Initialize config manager.
        config_name: Name of config file (without .json), e.g. 'OPPD'.
                     If None, uses active_config from config/_active.json.
        """
        self.config_name = config_name or get_active_config_name()
        self.config_file = self._find_config_file()
        self.config = self.load_config()
        self.state = _load_state()

    def _find_config_file(self) -> str:
        """Find config file in config/ folder."""
        config_dir = _get_config_dir()
        config_path = os.path.join(config_dir, f"{self.config_name}.json")

        if getattr(sys, "frozen", False):
            exe_dir = os.path.dirname(sys.executable)
            possible = [
                os.path.join(exe_dir, "config", f"{self.config_name}.json"),
                config_path,
                os.path.join(os.getcwd(), "config", f"{self.config_name}.json"),
            ]
            for p in possible:
                if os.path.exists(p):
                    return p

        if os.path.exists(config_path):
            return config_path

        os.makedirs(config_dir, exist_ok=True)
        return config_path

    def load_config(self) -> Dict[str, Any]:
        """Load static configuration from JSON file."""
        default_config = {
            "configurations": [
                {"name": "Default", "power_label": "POWER"},
                {"name": "OPPD", "power_label": "OPPD"},
            ],
            "selected_config": "OPPD",
        }

        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, "r", encoding="utf-8") as f:
                    config = json.load(f)
                for key, value in default_config.items():
                    if key not in config:
                        config[key] = value
                configs = config.get("configurations", [])
                if isinstance(configs, list):
                    names = [c.get("name") for c in configs if isinstance(c, dict)]
                    for req in [
                        {"name": "Default", "power_label": "POWER"},
                        {"name": "OPPD", "power_label": "OPPD"},
                    ]:
                        if req["name"] not in names:
                            configs = list(configs) + [req]
                            config["configurations"] = configs
                return config
        except Exception as e:
            print(f"Error loading config: {e}")

        return default_config

    def save_config(self) -> None:
        """Save static configuration to JSON file."""
        try:
            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=2)
        except Exception as e:
            print(f"Error saving config: {e}")

    def get(self, key: str, default: Any = None) -> Any:
        """Get configuration value (config or state)."""
        if key in STATE_KEYS:
            return self.state.get(key, default)
        return self.config.get(key, default)

    def set(self, key: str, value: Any) -> None:
        """Set configuration value and persist to appropriate file."""
        if key in STATE_KEYS:
            self.state[key] = value
            _save_state(self.state)
        else:
            self.config[key] = value
            self.save_config()

    def add_recent_file(self, file_path: str) -> None:
        """Add file to recent files list."""
        recent = self.state.get("recent_files", [])
        if file_path in recent:
            recent.remove(file_path)
        recent.insert(0, file_path)
        recent = recent[:10]
        self.set("recent_files", recent)

    def switch_config(self, name: str) -> None:
        """Switch to config from config/ folder and reload."""
        self.config_name = name
        self.config_file = self._find_config_file()
        self.config = self.load_config()
        set_active_config(name)

    def get_power_label(self) -> str:
        """Get power_label for current config."""
        label = self.config.get("power_label")
        if label:
            return label
        for c in self.config.get("configurations", []):
            if isinstance(c, dict) and c.get("name") == self.config_name:
                return c.get("power_label", "POWER")
        return "POWER"
