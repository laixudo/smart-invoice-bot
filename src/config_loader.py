from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml

from .paths import app_root, resource_root


def _deep_merge(base: dict, override: dict) -> dict:
    out = dict(base)
    for k, v in override.items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def load_config(path: str | Path | None = None) -> dict[str, Any]:
    root = app_root()
    if path:
        cfg_path = Path(path)
    else:
        cfg_path = root / "config" / "config.yaml"
        if not cfg_path.exists():
            # 单文件 exe：内置配置在 _MEIPASS
            cfg_path = resource_root() / "config" / "config.yaml"
    with cfg_path.open("r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    local_path = root / "config" / "config.local.yaml"
    if local_path.exists():
        with local_path.open("r", encoding="utf-8") as f:
            local = yaml.safe_load(f) or {}
        data = _deep_merge(data, local)
    return data
