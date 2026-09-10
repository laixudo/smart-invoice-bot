# -*- coding: utf-8 -*-
"""打包为单个 SMART发票助手.exe，同事双击即可用。"""
from pathlib import Path
import shutil
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]


def main():
    print("Installing PyInstaller...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "pyinstaller"])
    print("Building one-file exe...")
    subprocess.check_call(
        [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "SMART_Invoice_App.spec"],
        cwd=str(ROOT),
    )
    dist = ROOT / "dist"
    exe = dist / "SMART发票助手.exe"
    if not exe.exists():
        # 兼容编码显示差异，按模式查找
        cands = list(dist.glob("*.exe"))
        if not cands:
            raise SystemExit("未找到生成的 exe")
        exe = cands[0]

    # 附带一页说明（可选，exe 内也已嵌入）
    guide = ROOT / "使用说明.txt"
    if guide.exists():
        shutil.copy2(guide, dist / "使用说明.txt")

    # 清理旧的文件夹包，避免同事拿错
    for p in dist.iterdir():
        if p.is_dir() and p.name.startswith("SMART"):
            shutil.rmtree(p, ignore_errors=True)

    zip_path = dist / "SMART发票助手_单文件版.zip"
    if zip_path.exists():
        zip_path.unlink()
    # 打成小 zip：只含 exe + 说明（解压后双击 exe）
    import zipfile

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(exe, arcname=exe.name)
        readme = dist / "使用说明.txt"
        if readme.exists():
            zf.write(readme, arcname=readme.name)

    print("DONE.")
    print("单文件 exe:", exe)
    print("发给同事的 zip:", zip_path)
    print("说明：首次打开会稍慢（自解压），之后 input/output 会生成在 exe 同目录。")


if __name__ == "__main__":
    main()
