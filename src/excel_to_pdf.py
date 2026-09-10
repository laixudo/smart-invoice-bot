"""通过本机 Microsoft Excel 或 WPS 表格将 xlsx 导出为 PDF。

注意：不要在未安装 Excel 时盲目 Dispatch「Excel.Application」，
否则 Windows 可能弹出 Microsoft Store「获取 Excel」页面。
"""

from __future__ import annotations

import time
import winreg
from pathlib import Path

# progid 探测顺序
_EXCEL_PROGIDS = ("Excel.Application",)
_WPS_PROGIDS = ("Ket.Application", "et.Application")


def _reg_query(path: str, name: str | None = None) -> str | None:
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, path) as key:
            val, _ = winreg.QueryValueEx(key, name)
            return str(val) if val else None
    except OSError:
        return None


def _progid_server_path(progid: str) -> Path | None:
    """返回 COM LocalServer 可执行路径（存在才算真正可用）。"""
    clsid = _reg_query(rf"{progid}\CLSID")
    if not clsid:
        return None
    for sub in (
        rf"CLSID\{clsid}\LocalServer32",
        rf"Wow6432Node\CLSID\{clsid}\LocalServer32",
    ):
        raw = _reg_query(sub)
        if not raw:
            continue
        exe = raw.strip().strip('"').split(" ")[0].strip('"')
        p = Path(exe)
        if p.exists():
            # 排除会跳转微软商店的空壳（常见 WindowsApps 占位）
            low = str(p).lower()
            if "windowsapps" in low and "excel" in low:
                # 商店占位：视为不可用，避免强制打开商店
                continue
            return p
    return None


def excel_installed() -> bool:
    if any(_progid_server_path(p) for p in _EXCEL_PROGIDS):
        return True
    # App Paths 兜底
    try:
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe",
        ) as key:
            val, _ = winreg.QueryValueEx(key, None)
            return bool(val and Path(str(val)).exists())
    except OSError:
        return False


def wps_installed() -> bool:
    if any(_progid_server_path(p) for p in _WPS_PROGIDS):
        return True
    # 常见安装目录粗查
    roots = [
        Path(r"C:\Program Files\Kingsoft\WPS Office"),
        Path(r"C:\Program Files (x86)\Kingsoft\WPS Office"),
        Path(r"C:\Program Files\WPS Office"),
        Path(r"C:\Program Files (x86)\WPS Office"),
    ]
    for root in roots:
        if not root.exists():
            continue
        for p in root.rglob("et.exe"):
            return True
    return False


def list_office_choices() -> list[tuple[str, str]]:
    """
    返回 [(显示名, engine_key), ...]
    engine_key: auto | excel | wps
    """
    has_excel = excel_installed()
    has_wps = wps_installed()
    choices: list[tuple[str, str]] = []
    if has_excel or has_wps:
        hint = []
        if has_excel:
            hint.append("Excel")
        if has_wps:
            hint.append("WPS")
        choices.append((f"自动（已检测：{'/'.join(hint)}）", "auto"))
    if has_excel:
        choices.append(("Microsoft Excel", "excel"))
    if has_wps:
        choices.append(("WPS 表格", "wps"))
    return choices


def resolve_progid(engine: str = "auto") -> str:
    """根据用户选择解析可用 COM ProgID；都没有则抛错。"""
    engine = (engine or "auto").strip().lower()
    if engine == "excel":
        for p in _EXCEL_PROGIDS:
            if _progid_server_path(p):
                return p
        raise RuntimeError("未检测到 Microsoft Excel。请安装 Excel，或改选 WPS 表格。")
    if engine == "wps":
        for p in _WPS_PROGIDS:
            if _progid_server_path(p):
                return p
        raise RuntimeError("未检测到 WPS 表格。请安装 WPS，或改选 Microsoft Excel。")
    # auto：优先已安装的，WPS 与 Excel 都有时优先 Excel（版式更稳），仅有 WPS 则用 WPS
    if excel_installed():
        for p in _EXCEL_PROGIDS:
            if _progid_server_path(p):
                return p
    if wps_installed():
        for p in _WPS_PROGIDS:
            if _progid_server_path(p):
                return p
    raise RuntimeError(
        "未检测到 Microsoft Excel 或 WPS 表格。\n"
        "请安装其中之一后再生成（不会强制打开微软商店）。"
    )


class OfficeExporter:
    """复用同一个 Excel/WPS 进程批量导出，避免每个 PDF 都启动/退出（大幅提速）。"""

    def __init__(self, engine: str = "auto"):
        self.engine = engine
        self._app = None
        self._progid = ""

    def __enter__(self) -> "OfficeExporter":
        try:
            import win32com.client  # type: ignore
        except ImportError as e:
            raise RuntimeError("需要 pywin32：pip install pywin32") from e
        self._progid = resolve_progid(self.engine)
        self._app = win32com.client.DispatchEx(self._progid)
        try:
            self._app.Visible = False
        except Exception:
            pass
        try:
            self._app.DisplayAlerts = False
        except Exception:
            pass
        try:
            self._app.ScreenUpdating = False
        except Exception:
            pass
        return self

    def export(self, xlsx_path: Path, pdf_path: Path) -> Path:
        if self._app is None:
            raise RuntimeError("OfficeExporter 未启动")
        xlsx_path = Path(xlsx_path).resolve()
        pdf_path = Path(pdf_path).resolve()
        pdf_path.parent.mkdir(parents=True, exist_ok=True)
        if pdf_path.exists():
            pdf_path.unlink()
        wb = self._app.Workbooks.Open(str(xlsx_path))
        try:
            wb.ExportAsFixedFormat(0, str(pdf_path))
        finally:
            wb.Close(SaveChanges=False)
        if not pdf_path.exists():
            raise RuntimeError(f"导出 PDF 失败（引擎 {self._progid}）：{pdf_path}")
        return pdf_path

    def __exit__(self, exc_type, exc, tb) -> None:
        app = self._app
        self._app = None
        if app is None:
            return
        try:
            app.ScreenUpdating = True
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass
        time.sleep(0.05)


def excel_export_pdf(
    xlsx_path: Path,
    pdf_path: Path,
    engine: str = "auto",
    exporter: OfficeExporter | None = None,
) -> Path:
    """使用 Excel 或 WPS COM ExportAsFixedFormat 导出 PDF。"""
    if exporter is not None:
        return exporter.export(xlsx_path, pdf_path)

    with OfficeExporter(engine) as session:
        return session.export(xlsx_path, pdf_path)
