"""PyInstaller / 双击入口（无参数启动 GUI）。"""

from __future__ import annotations

import sys
import traceback
from pathlib import Path


def _crash_log_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent / "启动失败日志.txt"
    return Path(__file__).resolve().parent / "启动失败日志.txt"


def _write_crash_log(exc: BaseException) -> None:
    try:
        _crash_log_path().write_text(
            "".join(traceback.format_exception(type(exc), exc, exc.__traceback__)),
            encoding="utf-8",
        )
    except Exception:
        pass


if __name__ == "__main__":
    try:
        from src.gui import main

        main()
    except Exception as exc:
        _write_crash_log(exc)
        raise
