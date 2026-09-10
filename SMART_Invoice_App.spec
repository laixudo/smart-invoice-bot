# -*- mode: python ; coding: utf-8 -*-
# 单文件 exe：同事只需双击 SMART发票助手.exe
from PyInstaller.utils.hooks import collect_all

block_cipher = None

ctk_datas, ctk_binaries, ctk_hidden = collect_all("customtkinter")

a = Analysis(
    ['run_app.py'],
    pathex=[],
    binaries=ctk_binaries,
    datas=[
        ('config/config.yaml', 'config'),
        ('templates', 'templates'),
        ('使用说明.txt', '.'),
    ] + ctk_datas,
    hiddenimports=[
        'openpyxl',
        'reportlab',
        'yaml',
        'win32com',
        'win32com.client',
        'pythoncom',
        'pywintypes',
        'customtkinter',
        'darkdetect',
        'src',
        'src.gui',
        'src.pipeline',
        'src.excel_parser',
        'src.excel_template_fill',
        'src.excel_to_pdf',
        'src.pdf_renderer',
        'src.validator',
        'src.report_writer',
        'src.config_loader',
        'src.money',
        'src.models',
        'src.paths',
        'src.ui_motion',
        'src.local_settings',
        'src.mailer',
    ] + list(ctk_hidden),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pandas', 'numpy', 'pdfplumber', 'pdfminer'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='SMART发票助手',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
