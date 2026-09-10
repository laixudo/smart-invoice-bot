"""运行路径：开发目录 / PyInstaller 打包后均可用。"""

from __future__ import annotations

import sys
from pathlib import Path


def app_root() -> Path:
    """可写根目录：exe 同级（或源码项目根）。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def resource_root() -> Path:
    """只读资源（模板等）。优先 exe 同级，其次 PyInstaller 解压目录。"""
    root = app_root()
    if (root / "templates").exists():
        return root
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        meipass = Path(sys._MEIPASS)
        if (meipass / "templates").exists():
            return meipass
    return root
