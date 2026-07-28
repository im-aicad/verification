# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

from PyInstaller.utils.hooks import collect_submodules


server_dir = Path(SPECPATH)
project_dir = server_dir.parent
algorithm_dir = server_dir / "pycatia_wheelhouse_regulation_verification"

datas = [
    (str(project_dir / "web"), "web"),
    (str(algorithm_dir / "main.py"), "pycatia_wheelhouse_regulation_verification"),
    (str(algorithm_dir / "section_curve_export_tool.py"), "pycatia_wheelhouse_regulation_verification"),
    (str(algorithm_dir / "catia_picture_capture.py"), "pycatia_wheelhouse_regulation_verification"),
    (str(algorithm_dir / "catia_annotation_tools.py"), "pycatia_wheelhouse_regulation_verification"),
    (str(algorithm_dir / "requirements.md"), "pycatia_wheelhouse_regulation_verification"),
]

for catpart in algorithm_dir.glob("*.CATPart"):
    datas.append((str(catpart), "pycatia_wheelhouse_regulation_verification"))

hiddenimports = (
    collect_submodules("uvicorn")
    + collect_submodules("win32com")
    + collect_submodules("docx")
    + [
        "run_algorithm_worker",
        "pythoncom",
        "pywintypes",
        "win32timezone",
        "win32com.client",
        "multipart",
        "lxml",
        "lxml.etree",
    ]
)

a = Analysis(
    [str(server_dir / "server.py")],
    pathex=[str(server_dir)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="wheelhouse_server",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
