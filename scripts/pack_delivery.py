# -*- coding: utf-8 -*-
"""Assemble source + installer zip for delivery. Run after PyInstaller finishes."""
from __future__ import annotations

import hashlib
import shutil
import sys
import zipfile
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(Path(__file__).resolve().parent))
STAMP = date.today().strftime("%Y%m%d")
BUNDLE_NAME = f"SMART发票助手_最新版_{STAMP}"
OUT_ROOT = ROOT.parent / BUNDLE_NAME
ZIP_PATH = ROOT.parent / f"{BUNDLE_NAME}.zip"

SKIP_DIR_NAMES = {
    "__pycache__",
    ".venv",
    "venv",
    "_pack_venv311",
    "_build_deps",
    "_build_deps310",
    "build",
    "build_legacy_20260817",
    "build_pdfqa_20260817",
    "build_rebuilt",
    "build_rebuilt310",
    "dist",
    "tmp",
    "work",
    "_work",
    "logs",
    "output",
    "quarantine",
    "input",
    "release",
}
SKIP_FILE_NAMES = {
    "config.local.yaml",
    "user_prefs.yaml",
    "启动失败日志.txt",
}
SKIP_SUFFIXES = {".pyc", ".pyo"}


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def copy_source(dst: Path) -> None:
    keep_files = [
        "README.md",
        "requirements.txt",
        "run_app.py",
        "SMART_Invoice_App.spec",
        "使用说明.txt",
        ".gitignore",
    ]
    keep_dirs = ["src", "config", "templates", "samples", "tests", "scripts"]
    dst.mkdir(parents=True, exist_ok=True)
    for name in keep_files:
        src = ROOT / name
        if src.exists():
            shutil.copy2(src, dst / name)
    for name in keep_dirs:
        src = ROOT / name
        for p in src.rglob("*"):
            if p.is_dir():
                continue
            rel = p.relative_to(src)
            if any(part in SKIP_DIR_NAMES for part in rel.parts):
                continue
            if p.name in SKIP_FILE_NAMES:
                continue
            if p.suffix in SKIP_SUFFIXES:
                continue
            target = dst / name / rel
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(p, target)
    example = ROOT / "config" / "config.local.yaml.example"
    if example.exists():
        (dst / "config").mkdir(exist_ok=True)
        shutil.copy2(example, dst / "config" / "config.local.yaml.example")


def zip_dir(src: Path, zip_path: Path) -> None:
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in src.rglob("*"):
            if p.is_file():
                zf.write(p, arcname=str(Path(BUNDLE_NAME) / p.relative_to(src)))


def main() -> None:
    exe_cands = list((ROOT / "dist").glob("*.exe"))
    if not exe_cands:
        raise SystemExit("dist 下没有 exe，请先完成 PyInstaller")
    exe = exe_cands[0]
    if OUT_ROOT.exists():
        shutil.rmtree(OUT_ROOT)
    OUT_ROOT.mkdir(parents=True)

    # tutorial
    from gen_user_guide import build  # type: ignore

    doc, txt = build()
    doc.save(OUT_ROOT / "请先看_使用教程.docx")
    (OUT_ROOT / "请先看_使用教程.txt").write_text(txt, encoding="utf-8")

    inst = OUT_ROOT / "01_安装包"
    inst.mkdir()
    shutil.copy2(exe, inst / "SMART发票助手.exe")
    shutil.copy2(ROOT / "使用说明.txt", inst / "使用说明.txt")
    shutil.copy2(OUT_ROOT / "请先看_使用教程.docx", inst / "请先看_使用教程.docx")
    shutil.copy2(OUT_ROOT / "请先看_使用教程.txt", inst / "请先看_使用教程.txt")

    src_dir = OUT_ROOT / "02_源码"
    copy_source(src_dir)

    digest = sha256(inst / "SMART发票助手.exe")
    size_mb = (inst / "SMART发票助手.exe").stat().st_size / (1024 * 1024)
    version = (OUT_ROOT / "版本说明.txt")
    version.write_text(
        "\n".join(
            [
                "SMART 单票 Invoice 助手 — 交付包",
                f"版本：v0.3 新旧模板兼容完整复核版",
                f"打包日期：{date.today().isoformat()}",
                "Python：3.11（安装包 exe）",
                "形态：单文件 SMART发票助手.exe，同事无需安装 Python",
                f"安装包大小：约 {size_mb:.1f} MB",
                f"exe SHA256：{digest}",
                "",
                "给业务同事：解压后打开 01_安装包，双击 SMART发票助手.exe",
                "给技术人员：02_源码 可开发与重新打包",
                "",
                "本包已排除：邮箱授权码、构建缓存、客户真实账单、虚拟环境",
                "本期不做：自动发送邮件",
                "",
            ]
        ),
        encoding="utf-8",
    )

    zip_dir(OUT_ROOT, ZIP_PATH)
    print("BUNDLE", OUT_ROOT)
    print("ZIP", ZIP_PATH, ZIP_PATH.stat().st_size)
    print("EXE", inst / "SMART发票助手.exe", f"{size_mb:.1f}MB")
    print("SHA256", digest)


if __name__ == "__main__":
    main()
