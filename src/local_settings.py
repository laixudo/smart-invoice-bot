"""本地用户偏好与本地敏感配置。"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

import yaml

from .paths import app_root


def _prefs_dir() -> Path:
    """偏好存用户目录，避免 exe 旁自动冒出文件夹。"""
    base = os.environ.get("APPDATA") or os.environ.get("LOCALAPPDATA")
    if base:
        d = Path(base) / "SMART发票助手"
    else:
        d = Path.home() / ".smart_invoice_bot"
    return d


def prefs_path() -> Path:
    return _prefs_dir() / "user_prefs.yaml"


def local_config_path() -> Path:
    # 如需本地覆盖，仍可用 exe 旁 config（不自动创建）
    return app_root() / "config" / "config.local.yaml"


def load_yaml(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    with path.open("r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def save_yaml(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        yaml.safe_dump(data, f, allow_unicode=True, sort_keys=False)


def load_prefs() -> dict[str, Any]:
    return load_yaml(prefs_path())


def save_prefs(data: dict[str, Any]) -> None:
    save_yaml(prefs_path(), data)


def load_local_email() -> dict[str, Any]:
    """兼容旧发信模块；本期 GUI 不发信。"""
    data = load_yaml(local_config_path())
    return data.get("email") or data
